import os
import io
import math
import logging
import asyncio
import asyncpg
from aiohttp import web
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import CommandStart
from aiogram.types import (
    Message,
    CallbackQuery,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    BufferedInputFile
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

# --- НАСТРОЙКИ ---
BOT_TOKEN = os.getenv("BOT_TOKEN")
DATABASE_URL = os.getenv("DATABASE_URL")
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))
PORT = int(os.getenv("PORT", "8080"))
PAGE_SIZE = 8

logging.basicConfig(level=logging.INFO)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
router = Router()
dp.include_router(router)

db_pool: asyncpg.Pool = None

# --- FSM СОСТОЯНИЯ ---
class AdminStates(StatesGroup):
    waiting_for_list_title = State()
    waiting_for_topics = State()
    waiting_for_append_topics = State()
    waiting_for_new_topic_title = State()
    waiting_for_new_max_members = State()

class StudentStates(StatesGroup):
    waiting_for_custom_topic = State()
    waiting_for_search_query = State()

# --- ИНИЦИАЛИЗАЦИЯ БАЗЫ ---
async def init_db():
    global db_pool
    db_pool = await asyncpg.create_pool(
        dsn=DATABASE_URL,
        min_size=2,
        max_size=10,
        command_timeout=15
    )
    async with db_pool.acquire() as conn:
        await conn.execute("""
            CREATE TABLE IF NOT EXISTS lists (
                id SERIAL PRIMARY KEY,
                title TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS topics (
                id SERIAL PRIMARY KEY,
                list_id INTEGER REFERENCES lists(id) ON DELETE CASCADE,
                title TEXT NOT NULL,
                max_members INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS topic_members (
                id SERIAL PRIMARY KEY,
                topic_id INTEGER REFERENCES topics(id) ON DELETE CASCADE,
                user_id BIGINT NOT NULL,
                user_name TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_topics_list ON topics(list_id);
            CREATE INDEX IF NOT EXISTS idx_topic_members_topic ON topic_members(topic_id);
            CREATE INDEX IF NOT EXISTS idx_topic_members_user ON topic_members(user_id);
        """)

# --- АСИНХРОННЫЕ ФУНКЦИИ БАЗЫ ДАННЫХ ---
async def get_all_lists():
    async with db_pool.acquire() as conn:
        return await conn.fetch("SELECT id, title FROM lists ORDER BY id DESC")

async def get_list_by_id(list_id):
    async with db_pool.acquire() as conn:
        return await conn.fetchrow("SELECT id, title FROM lists WHERE id = $1", list_id)

async def get_topics(list_id, search_query=None):
    async with db_pool.acquire() as conn:
        if search_query:
            return await conn.fetch(
                "SELECT id, title, max_members FROM topics WHERE list_id = $1 AND title ILIKE $2 ORDER BY id ASC",
                list_id, f"%{search_query}%"
            )
        return await conn.fetch(
            "SELECT id, title, max_members FROM topics WHERE list_id = $1 ORDER BY id ASC",
            list_id
        )

async def get_topic_by_id(topic_id):
    async with db_pool.acquire() as conn:
        return await conn.fetchrow("SELECT id, list_id, title, max_members FROM topics WHERE id = $1", topic_id)

async def get_members(topic_id):
    async with db_pool.acquire() as conn:
        return await conn.fetch("SELECT user_id, user_name FROM topic_members WHERE topic_id = $1", topic_id)

async def get_user_topic_in_list(list_id, user_id):
    async with db_pool.acquire() as conn:
        return await conn.fetchrow("""
            SELECT t.id, t.title 
            FROM topic_members tm
            JOIN topics t ON tm.topic_id = t.id
            WHERE t.list_id = $1 AND tm.user_id = $2
        """, list_id, user_id)

async def add_member(topic_id, user_id, user_name):
    async with db_pool.acquire() as conn:
        async with conn.transaction():
            res = await conn.fetchrow("SELECT list_id, max_members FROM topics WHERE id = $1", topic_id)
            if not res:
                return False, "Тема не найдена."
            list_id, max_m = res["list_id"], res["max_members"]

            existing = await conn.fetchrow("""
                SELECT t.title 
                FROM topic_members tm
                JOIN topics t ON tm.topic_id = t.id
                WHERE t.list_id = $1 AND tm.user_id = $2
            """, list_id, user_id)
            if existing:
                return False, f"Вы уже записаны на тему «{existing['title']}». Можно выбрать только одну тему!"

            cur_m = await conn.fetchval("SELECT COUNT(*) FROM topic_members WHERE topic_id = $1", topic_id)
            if cur_m >= max_m:
                return False, "Места на эту тему уже закончились!"

            await conn.execute("INSERT INTO topic_members (topic_id, user_id, user_name) VALUES ($1, $2, $3)", topic_id, user_id, user_name)
            return True, "Успешно!"

async def remove_member(topic_id, user_id):
    async with db_pool.acquire() as conn:
        await conn.execute("DELETE FROM topic_members WHERE topic_id = $1 AND user_id = $2", topic_id, user_id)

async def clear_topic_members(topic_id):
    async with db_pool.acquire() as conn:
        await conn.execute("DELETE FROM topic_members WHERE topic_id = $1", topic_id)

async def get_user_topics(user_id):
    async with db_pool.acquire() as conn:
        return await conn.fetch("""
            SELECT t.id, t.title, l.title AS list_title
            FROM topic_members tm
            JOIN topics t ON tm.topic_id = t.id
            JOIN lists l ON t.list_id = l.id
            WHERE tm.user_id = $1
        """, user_id)

async def create_list(title):
    async with db_pool.acquire() as conn:
        return await conn.fetchval("INSERT INTO lists (title) VALUES ($1) RETURNING id", title)

async def create_topic(list_id, title, max_members=1):
    async with db_pool.acquire() as conn:
        return await conn.fetchval("INSERT INTO topics (list_id, title, max_members) VALUES ($1, $2, $3) RETURNING id", list_id, title, max_members)

async def update_topic_title(topic_id, new_title):
    async with db_pool.acquire() as conn:
        await conn.execute("UPDATE topics SET title = $1 WHERE id = $2", new_title, topic_id)

async def update_topic_max_members(topic_id, new_max):
    async with db_pool.acquire() as conn:
        await conn.execute("UPDATE topics SET max_members = $1 WHERE id = $2", new_max, topic_id)

async def delete_topic(topic_id):
    async with db_pool.acquire() as conn:
        await conn.execute("DELETE FROM topics WHERE id = $1", topic_id)

async def batch_create_topics(topics_data):
    async with db_pool.acquire() as conn:
        await conn.executemany("INSERT INTO topics (list_id, title, max_members) VALUES ($1, $2, $3)", topics_data)

async def delete_list(list_id):
    async with db_pool.acquire() as conn:
        await conn.execute("DELETE FROM lists WHERE id = $1", list_id)

# --- КЛАВИАТУРЫ ---
def get_main_menu(is_admin=False):
    kb = [
        [InlineKeyboardButton(text="📋 Выбрать предмет и тему", callback_data="view_lists")],
        [InlineKeyboardButton(text="👤 Мои выбранные темы", callback_data="my_topics")],
        [InlineKeyboardButton(text="📥 Скачать реестр (Excel)", callback_data="download_excel")]
    ]
    if is_admin:
        kb.append([InlineKeyboardButton(text="⚙️ Меню старосты", callback_data="admin_menu")])
    return InlineKeyboardMarkup(inline_keyboard=kb)

def get_admin_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="➕ Создать новый предмет", callback_data="admin_create_list")],
        [InlineKeyboardButton(text="📥 Добавить темы в предмет", callback_data="admin_append_topics_choose")],
        [InlineKeyboardButton(text="✏️ Редактировать темы", callback_data="admin_edit_topics_choose")],
        [InlineKeyboardButton(text="🗑 Удалить предмет со всеми темами", callback_data="admin_delete_list")],
        [InlineKeyboardButton(text="◀️ Главное меню", callback_data="main_menu")]
    ])

async def build_topics_page_keyboard(list_id, topics, user_id, page=0, is_search=False, is_admin=False):
    total_topics = len(topics)
    total_pages = max(1, math.ceil(total_topics / PAGE_SIZE))
    page = max(0, min(page, total_pages - 1))

    start_idx = page * PAGE_SIZE
    end_idx = start_idx + PAGE_SIZE
    current_page_topics = topics[start_idx:end_idx]

    kb = []
    for topic in current_page_topics:
        tid, title, max_m = topic["id"], topic["title"], topic["max_members"]
        members = await get_members(tid)
        cur_m = len(members)
        is_my_topic = any(m["user_id"] == user_id for m in members)

        if is_my_topic:
            status = "⭐️ ВЫ"
        elif cur_m >= max_m:
            status = "🔴 Занято"
        else:
            status = f"🟢 ({cur_m}/{max_m})"

        display_title = title if len(title) <= 35 else title[:32] + "..."
        btn_text = f"{display_title} — {status}"
        kb.append([InlineKeyboardButton(text=btn_text, callback_data=f"topic_{tid}_{page}")])

    nav_row = []
    if page > 0:
        nav_row.append(InlineKeyboardButton(text="⬅️ Назад", callback_data=f"page_{list_id}_{page - 1}"))
    if total_pages > 1:
        nav_row.append(InlineKeyboardButton(text=f"Стр. {page + 1}/{total_pages}", callback_data="noop"))
    if page < total_pages - 1:
        nav_row.append(InlineKeyboardButton(text="Вперёд ➡️", callback_data=f"page_{list_id}_{page + 1}"))

    if nav_row:
        kb.append(nav_row)

    bottom_row = []
    if is_search:
        bottom_row.append(InlineKeyboardButton(text="🔄 Сбросить поиск", callback_data=f"open_list_{list_id}_0"))
    else:
        bottom_row.append(InlineKeyboardButton(text="🔍 Поиск темы", callback_data=f"search_topic_{list_id}"))

    user_chosen = await get_user_topic_in_list(list_id, user_id)
    if not user_chosen:
        bottom_row.append(InlineKeyboardButton(text="💡 «Другое»", callback_data=f"custom_topic_{list_id}"))

    if bottom_row:
        kb.append(bottom_row)

    if is_admin:
        kb.append([InlineKeyboardButton(text="➕ Добавить темы в этот список", callback_data=f"admin_append_to_{list_id}")])

    kb.append([InlineKeyboardButton(text="◀️ К списку предметов", callback_data="view_lists")])
    return InlineKeyboardMarkup(inline_keyboard=kb)

async def build_admin_topics_page_keyboard(list_id, topics, page=0):
    total_topics = len(topics)
    total_pages = max(1, math.ceil(total_topics / PAGE_SIZE))
    page = max(0, min(page, total_pages - 1))

    start_idx = page * PAGE_SIZE
    end_idx = start_idx + PAGE_SIZE
    current_page_topics = topics[start_idx:end_idx]

    kb = []
    for topic in current_page_topics:
        tid, title, max_m = topic["id"], topic["title"], topic["max_members"]
        members = await get_members(tid)
        cur_m = len(members)
        display_title = title if len(title) <= 32 else title[:29] + "..."
        btn_text = f"✏️ {display_title} ({cur_m}/{max_m})"
        kb.append([InlineKeyboardButton(text=btn_text, callback_data=f"admin_topic_{tid}_{page}")])

    nav_row = []
    if page > 0:
        nav_row.append(InlineKeyboardButton(text="⬅️ Назад", callback_data=f"admin_page_{list_id}_{page - 1}"))
    if total_pages > 1:
        nav_row.append(InlineKeyboardButton(text=f"Стр. {page + 1}/{total_pages}", callback_data="noop"))
    if page < total_pages - 1:
        nav_row.append(InlineKeyboardButton(text="Вперёд ➡️", callback_data=f"admin_page_{list_id}_{page + 1}"))

    if nav_row:
        kb.append(nav_row)

    kb.append([InlineKeyboardButton(text="➕ Добавить темы в этот предмет", callback_data=f"admin_append_to_{list_id}")])
    kb.append([InlineKeyboardButton(text="◀️ К выбору предметов", callback_data="admin_edit_topics_choose")])
    return InlineKeyboardMarkup(inline_keyboard=kb)

# --- ХЕНДЛЕРЫ ---
@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    is_admin = (ADMIN_ID == 0) or (message.from_user.id == ADMIN_ID)
    await message.answer(
        f"👋 Привет, {message.from_user.first_name}!\n\nГлавное меню каталога тем:",
        reply_markup=get_main_menu(is_admin)
    )

@router.callback_query(F.data == "main_menu")
async def cb_main_menu(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    is_admin = (ADMIN_ID == 0) or (callback.from_user.id == ADMIN_ID)
    await callback.message.edit_text("Главное меню каталога тем:", reply_markup=get_main_menu(is_admin))

@router.callback_query(F.data == "noop")
async def cb_noop(callback: CallbackQuery):
    await callback.answer()

@router.callback_query(F.data == "admin_menu")
async def cb_admin_menu(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    if ADMIN_ID != 0 and callback.from_user.id != ADMIN_ID:
        await callback.answer("У вас нет прав старосты.", show_alert=True)
        return
    await callback.message.edit_text("⚙️ <b>Панель старосты:</b>\nВыберите действие:", parse_mode="HTML", reply_markup=get_admin_menu())

@router.callback_query(F.data == "admin_create_list")
async def cb_admin_create_list(callback: CallbackQuery, state: FSMContext):
    await state.set_state(AdminStates.waiting_for_list_title)
    await callback.message.edit_text("✏️ Введите название предмета (например: <i>Макроэкономика</i>):", parse_mode="HTML")

@router.message(AdminStates.waiting_for_list_title)
async def process_list_title(message: Message, state: FSMContext):
    title = message.text.strip()
    list_id = await create_list(title)
    await state.update_data(list_id=list_id, list_title=title)
    await state.set_state(AdminStates.waiting_for_topics)
    await message.answer(
        f"Предмет <b>«{title}»</b> создан!\n\nОтправьте список тем сообщением (каждая с новой строки).\n<i>Если нужно несколько мест: Тема (2)</i>",
        parse_mode="HTML"
    )

@router.message(AdminStates.waiting_for_topics)
async def process_topics(message: Message, state: FSMContext):
    data = await state.get_data()
    list_id = data.get("list_id")
    lines = [line.strip() for line in message.text.split("\n") if line.strip()]

    topics_to_insert = []
    for line in lines:
        max_m = 1
        title = line
        if line.endswith(")") and "(" in line:
            try:
                num = line.rsplit("(", 1)[1].rstrip(")")
                max_m = int(num)
                title = line.rsplit("(", 1)[0].strip()
            except ValueError:
                pass
        topics_to_insert.append((list_id, title, max_m))

    if topics_to_insert:
        await batch_create_topics(topics_to_insert)

    await state.clear()
    is_admin = (ADMIN_ID == 0) or (message.from_user.id == ADMIN_ID)
    await message.answer(
        f"✅ Успешно добавлено {len(topics_to_insert)} тем к предмету <b>«{data.get('list_title')}»</b>!",
        parse_mode="HTML",
        reply_markup=get_main_menu(is_admin)
    )

@router.callback_query(F.data == "admin_append_topics_choose")
async def cb_admin_append_topics_choose(callback: CallbackQuery):
    lists = await get_all_lists()
    if not lists:
        await callback.message.edit_text("Нет созданных предметов.", reply_markup=get_admin_menu())
        return
    kb = [[InlineKeyboardButton(text=t["title"], callback_data=f"admin_append_to_{t['id']}")] for t in lists]
    kb.append([InlineKeyboardButton(text="◀️ Назад", callback_data="admin_menu")])
    await callback.message.edit_text("Выберите предмет, в который нужно добавить темы:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("admin_append_to_"))
async def cb_admin_append_to(callback: CallbackQuery, state: FSMContext):
    list_id = int(callback.data.split("_")[3])
    item = await get_list_by_id(list_id)
    await state.update_data(append_list_id=list_id, append_list_title=item["title"])
    await state.set_state(AdminStates.waiting_for_append_topics)
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Отмена", callback_data="admin_menu")]])
    await callback.message.edit_text(
        f"📥 Отправьте новые темы для предмета <b>«{item['title']}»</b> (каждая с новой строки):",
        parse_mode="HTML",
        reply_markup=kb
    )

@router.message(AdminStates.waiting_for_append_topics)
async def process_append_topics(message: Message, state: FSMContext):
    data = await state.get_data()
    list_id = data.get("append_list_id")
    lines = [line.strip() for line in message.text.split("\n") if line.strip()]

    topics_to_insert = []
    for line in lines:
        max_m = 1
        title = line
        if line.endswith(")") and "(" in line:
            try:
                num = line.rsplit("(", 1)[1].rstrip(")")
                max_m = int(num)
                title = line.rsplit("(", 1)[0].strip()
            except ValueError:
                pass
        topics_to_insert.append((list_id, title, max_m))

    if topics_to_insert:
        await batch_create_topics(topics_to_insert)

    await state.clear()
    is_admin = (ADMIN_ID == 0) or (message.from_user.id == ADMIN_ID)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📋 Открыть этот предмет", callback_data=f"open_list_{list_id}_0")],
        [InlineKeyboardButton(text="◀️ Меню старосты", callback_data="admin_menu")]
    ])
    await message.answer(
        f"✅ Добавлено {len(topics_to_insert)} новых тем к предмету <b>«{data.get('append_list_title')}»</b>!",
        parse_mode="HTML",
        reply_markup=kb
    )

@router.callback_query(F.data == "admin_edit_topics_choose")
async def cb_admin_edit_topics_choose(callback: CallbackQuery):
    lists = await get_all_lists()
    if not lists:
        await callback.message.edit_text("Нет созданных предметов.", reply_markup=get_admin_menu())
        return
    kb = [[InlineKeyboardButton(text=t["title"], callback_data=f"admin_open_edit_list_{t['id']}_0")] for t in lists]
    kb.append([InlineKeyboardButton(text="◀️ Назад", callback_data="admin_menu")])
    await callback.message.edit_text("Выберите предмет для редактирования тем:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("admin_open_edit_list_"))
async def cb_admin_open_edit_list(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    parts = callback.data.split("_")
    list_id = int(parts[4])
    page = int(parts[5]) if len(parts) > 5 else 0

    topics = await get_topics(list_id)
    item = await get_list_by_id(list_id)
    if not topics:
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="➕ Добавить темы в этот предмет", callback_data=f"admin_append_to_{list_id}")],
            [InlineKeyboardButton(text="◀️ Назад", callback_data="admin_edit_topics_choose")]
        ])
        await callback.message.edit_text(f"В предмете «{item['title']}» пока нет тем.", reply_markup=kb)
        return

    kb = await build_admin_topics_page_keyboard(list_id, topics, page=page)
    await callback.message.edit_text(
        f"✏️ <b>Редактирование тем: «{item['title']}»</b>\nВыберите тему:",
        parse_mode="HTML",
        reply_markup=kb
    )

@router.callback_query(F.data.startswith("admin_page_"))
async def cb_admin_page(callback: CallbackQuery):
    parts = callback.data.split("_")
    list_id = int(parts[2])
    page = int(parts[3])

    topics = await get_topics(list_id)
    item = await get_list_by_id(list_id)
    kb = await build_admin_topics_page_keyboard(list_id, topics, page=page)
    await callback.message.edit_text(
        f"✏️ <b>Редактирование тем: «{item['title']}»</b>\nВыберите тему:",
        parse_mode="HTML",
        reply_markup=kb
    )

@router.callback_query(F.data.startswith("admin_topic_"))
async def cb_admin_topic_detail(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    parts = callback.data.split("_")
    topic_id = int(parts[2])
    page = int(parts[3]) if len(parts) > 3 else 0

    topic = await get_topic_by_id(topic_id)
    if not topic:
        await callback.answer("Тема не найдена")
        return

    list_id, title, max_m = topic["list_id"], topic["title"], topic["max_members"]
    members = await get_members(topic_id)

    text = f"⚙️ <b>Управление темой:</b>\n\n<b>Название:</b> {title}\n<b>Лимит мест:</b> {len(members)}/{max_m}\n\n<b>Записаны:</b>\n"
    for m in members:
        text += f"• {m['user_name']}\n"
    if not members:
        text += "— Свободно\n"

    kb = [
        [InlineKeyboardButton(text="✏️ Изменить название", callback_data=f"edit_title_{topic_id}_{page}")],
        [InlineKeyboardButton(text="👥 Изменить лимит мест", callback_data=f"edit_limit_{topic_id}_{page}")],
        [InlineKeyboardButton(text="🔄 Снять всех студентов", callback_data=f"clear_mem_{topic_id}_{page}")],
        [InlineKeyboardButton(text="🗑 Удалить тему полностью", callback_data=f"del_topic_{topic_id}_{page}")],
        [InlineKeyboardButton(text="◀️ Назад к темам", callback_data=f"admin_open_edit_list_{list_id}_{page}")]
    ]
    await callback.message.edit_text(text, parse_mode="HTML", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("edit_title_"))
async def cb_edit_title(callback: CallbackQuery, state: FSMContext):
    parts = callback.data.split("_")
    topic_id = int(parts[2])
    page = int(parts[3])

    await state.update_data(edit_topic_id=topic_id, edit_topic_page=page)
    await state.set_state(AdminStates.waiting_for_new_topic_title)
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Отмена", callback_data=f"admin_topic_{topic_id}_{page}")]])
    await callback.message.edit_text("Отправьте новое название темы в чат:", reply_markup=kb)

@router.message(AdminStates.waiting_for_new_topic_title)
async def process_new_topic_title(message: Message, state: FSMContext):
    new_title = message.text.strip()
    data = await state.get_data()
    topic_id = data.get("edit_topic_id")
    page = data.get("edit_topic_page")

    await update_topic_title(topic_id, new_title)
    await state.clear()
    await message.answer("✅ Название темы обновлено!")

    topic = await get_topic_by_id(topic_id)
    list_id, title, max_m = topic["list_id"], topic["title"], topic["max_members"]
    members = await get_members(topic_id)
    text = f"⚙️ <b>Управление темой:</b>\n\n<b>Название:</b> {title}\n<b>Лимит мест:</b> {len(members)}/{max_m}\n\n<b>Записаны:</b>\n"
    for m in members:
        text += f"• {m['user_name']}\n"
    if not members:
        text += "— Свободно\n"

    kb = [
        [InlineKeyboardButton(text="✏️ Изменить название", callback_data=f"edit_title_{topic_id}_{page}")],
        [InlineKeyboardButton(text="👥 Изменить лимит мест", callback_data=f"edit_limit_{topic_id}_{page}")],
        [InlineKeyboardButton(text="🔄 Снять всех студентов", callback_data=f"clear_mem_{topic_id}_{page}")],
        [InlineKeyboardButton(text="🗑 Удалить тему полностью", callback_data=f"del_topic_{topic_id}_{page}")],
        [InlineKeyboardButton(text="◀️ Назад к темам", callback_data=f"admin_open_edit_list_{list_id}_{page}")]
    ]
    await message.answer(text, parse_mode="HTML", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("edit_limit_"))
async def cb_edit_limit(callback: CallbackQuery, state: FSMContext):
    parts = callback.data.split("_")
    topic_id = int(parts[2])
    page = int(parts[3])

    await state.update_data(edit_topic_id=topic_id, edit_topic_page=page)
    await state.set_state(AdminStates.waiting_for_new_max_members)
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Отмена", callback_data=f"admin_topic_{topic_id}_{page}")]])
    await callback.message.edit_text("Отправьте новое максимальное количество мест (число от 1 до 10):", reply_markup=kb)

@router.message(AdminStates.waiting_for_new_max_members)
async def process_new_max_members(message: Message, state: FSMContext):
    data = await state.get_data()
    topic_id = data.get("edit_topic_id")
    page = data.get("edit_topic_page")

    try:
        new_max = int(message.text.strip())
        if new_max < 1:
            raise ValueError
    except ValueError:
        await message.answer("Пожалуйста, введите корректное число от 1 до 10.")
        return

    await update_topic_max_members(topic_id, new_max)
    await state.clear()
    await message.answer(f"✅ Лимит мест изменен на {new_max}!")

    topic = await get_topic_by_id(topic_id)
    list_id, title, max_m = topic["list_id"], topic["title"], topic["max_members"]
    members = await get_members(topic_id)
    text = f"⚙️ <b>Управление темой:</b>\n\n<b>Название:</b> {title}\n<b>Лимит мест:</b> {len(members)}/{max_m}\n\n<b>Записаны:</b>\n"
    for m in members:
        text += f"• {m['user_name']}\n"
    if not members:
        text += "— Свободно\n"

    kb = [
        [InlineKeyboardButton(text="✏️ Изменить название", callback_data=f"edit_title_{topic_id}_{page}")],
        [InlineKeyboardButton(text="👥 Изменить лимит мест", callback_data=f"edit_limit_{topic_id}_{page}")],
        [InlineKeyboardButton(text="🔄 Снять всех студентов", callback_data=f"clear_mem_{topic_id}_{page}")],
        [InlineKeyboardButton(text="🗑 Удалить тему полностью", callback_data=f"del_topic_{topic_id}_{page}")],
        [InlineKeyboardButton(text="◀️ Назад к темам", callback_data=f"admin_open_edit_list_{list_id}_{page}")]
    ]
    await message.answer(text, parse_mode="HTML", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("clear_mem_"))
async def cb_clear_mem(callback: CallbackQuery):
    parts = callback.data.split("_")
    topic_id = int(parts[2])
    page = int(parts[3])
    await clear_topic_members(topic_id)
    await callback.answer("Все студенты сняты с темы!")
    await cb_admin_topic_detail(callback, None)

@router.callback_query(F.data.startswith("del_topic_"))
async def cb_del_topic(callback: CallbackQuery):
    parts = callback.data.split("_")
    topic_id = int(parts[2])
    page = int(parts[3])

    topic = await get_topic_by_id(topic_id)
    list_id = topic["list_id"]
    await delete_topic(topic_id)
    await callback.answer("Тема удалена!")

    topics = await get_topics(list_id)
    item = await get_list_by_id(list_id)
    kb = await build_admin_topics_page_keyboard(list_id, topics, page=page)
    await callback.message.edit_text(
        f"✏️ <b>Редактирование тем: «{item['title']}»</b>\nТема удалена:",
        parse_mode="HTML",
        reply_markup=kb
    )

@router.callback_query(F.data == "admin_delete_list")
async def cb_admin_delete_list(callback: CallbackQuery):
    lists = await get_all_lists()
    if not lists:
        await callback.message.edit_text("Нет доступных предметов для удаления.", reply_markup=get_admin_menu())
        return
    kb = [[InlineKeyboardButton(text=f"❌ {t['title']}", callback_data=f"del_list_{t['id']}")] for t in lists]
    kb.append([InlineKeyboardButton(text="◀️ Назад", callback_data="admin_menu")])
    await callback.message.edit_text("Выберите предмет, который хотите полностью удалить:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("del_list_"))
async def cb_confirm_del_list(callback: CallbackQuery):
    list_id = int(callback.data.split("_")[2])
    await delete_list(list_id)
    await callback.answer("Предмет и все темы удалены!")
    await cb_admin_delete_list(callback)

# --- ПРОСМОТР ТЕМ И ПАГИНАЦИЯ ---
@router.callback_query(F.data == "view_lists")
async def cb_view_lists(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    lists = await get_all_lists()
    if not lists:
        back_kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Назад", callback_data="main_menu")]])
        await callback.message.edit_text("Пока нет доступных предметов.", reply_markup=back_kb)
        return
    kb = [[InlineKeyboardButton(text=t["title"], callback_data=f"open_list_{t['id']}_0")] for t in lists]
    kb.append([InlineKeyboardButton(text="◀️ Главное меню", callback_data="main_menu")])
    await callback.message.edit_text("Выберите предмет:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("open_list_"))
async def cb_open_list(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    parts = callback.data.split("_")
    list_id = int(parts[2])
    page = int(parts[3]) if len(parts) > 3 else 0

    topics = await get_topics(list_id)
    is_admin = (ADMIN_ID == 0) or (callback.from_user.id == ADMIN_ID)

    if not topics:
        kb_rows = [[InlineKeyboardButton(text="💡 Предложить свою тему («Другое»)", callback_data=f"custom_topic_{list_id}")]]
        if is_admin:
            kb_rows.append([InlineKeyboardButton(text="➕ Добавить темы в этот список", callback_data=f"admin_append_to_{list_id}")])
        kb_rows.append([InlineKeyboardButton(text="◀️ Назад", callback_data="view_lists")])
        await callback.message.edit_text("В этом предмете пока нет тем. Вы можете предложить тему:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb_rows))
        return

    user_chosen = await get_user_topic_in_list(list_id, callback.from_user.id)
    header_text = f"📋 <b>Каталог тем</b> (Всего тем: {len(topics)})\n"
    if user_chosen:
        header_text += f"\n⚠️ <i>Вы уже записаны на тему: «{user_chosen['title']}»</i>"

    kb = await build_topics_page_keyboard(list_id, topics, callback.from_user.id, page=page, is_admin=is_admin)
    await callback.message.edit_text(header_text, parse_mode="HTML", reply_markup=kb)

@router.callback_query(F.data.startswith("page_"))
async def cb_change_page(callback: CallbackQuery):
    parts = callback.data.split("_")
    list_id = int(parts[1])
    page = int(parts[2])

    topics = await get_topics(list_id)
    is_admin = (ADMIN_ID == 0) or (callback.from_user.id == ADMIN_ID)
    user_chosen = await get_user_topic_in_list(list_id, callback.from_user.id)
    header_text = f"📋 <b>Каталог тем</b> (Всего тем: {len(topics)})\n"
    if user_chosen:
        header_text += f"\n⚠️ <i>Вы уже записаны на тему: «{user_chosen['title']}»</i>"

    kb = await build_topics_page_keyboard(list_id, topics, callback.from_user.id, page=page, is_admin=is_admin)
    await callback.message.edit_text(header_text, parse_mode="HTML", reply_markup=kb)

# Поиск
@router.callback_query(F.data.startswith("search_topic_"))
async def cb_search_topic(callback: CallbackQuery, state: FSMContext):
    list_id = int(callback.data.split("_")[2])
    await state.update_data(search_list_id=list_id)
    await state.set_state(StudentStates.waiting_for_search_query)
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Отмена", callback_data=f"open_list_{list_id}_0")]])
    await callback.message.edit_text("🔍 <b>Поиск по темам</b>\n\nОтправьте ключевое слово или часть названия темы:", parse_mode="HTML", reply_markup=kb)

@router.message(StudentStates.waiting_for_search_query)
async def process_search_query(message: Message, state: FSMContext):
    query = message.text.strip()
    data = await state.get_data()
    list_id = data.get("search_list_id")
    await state.clear()

    topics = await get_topics(list_id, search_query=query)
    if not topics:
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🔄 Сбросить поиск", callback_data=f"open_list_{list_id}_0")],
            [InlineKeyboardButton(text="🔍 Искать снова", callback_data=f"search_topic_{list_id}")]
        ])
        await message.answer(f"По запросу «{query}» ничего не найдено.", reply_markup=kb)
        return

    is_admin = (ADMIN_ID == 0) or (message.from_user.id == ADMIN_ID)
    kb = await build_topics_page_keyboard(list_id, topics, message.from_user.id, page=0, is_search=True, is_admin=is_admin)
    await message.answer(f"🔍 Результаты поиска по запросу «<b>{query}</b>» (найдено: {len(topics)}):", parse_mode="HTML", reply_markup=kb)

# Предложить тему («Другое»)
@router.callback_query(F.data.startswith("custom_topic_"))
async def cb_custom_topic_start(callback: CallbackQuery, state: FSMContext):
    list_id = int(callback.data.split("_")[2])
    user_chosen = await get_user_topic_in_list(list_id, callback.from_user.id)
    if user_chosen:
        await callback.answer("Вы уже записаны на тему в этом предмете!", show_alert=True)
        return

    await state.update_data(custom_list_id=list_id)
    await state.set_state(StudentStates.waiting_for_custom_topic)
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Отмена", callback_data=f"open_list_{list_id}_0")]])
    await callback.message.edit_text("💡 <b>Предложить свою тему («Другое»)</b>\n\nНапишите точное название темы сообщением в чат:", parse_mode="HTML", reply_markup=kb)

@router.message(StudentStates.waiting_for_custom_topic)
async def process_custom_topic_text(message: Message, state: FSMContext):
    topic_title = message.text.strip()
    if not topic_title:
        await message.answer("Пожалуйста, отправьте непустой текст темы.")
        return

    data = await state.get_data()
    list_id = data.get("custom_list_id")

    user_chosen = await get_user_topic_in_list(list_id, message.from_user.id)
    if user_chosen:
        await state.clear()
        await message.answer("Вы уже записаны на тему в этом предмете.")
        return

    topic_id = await create_topic(list_id, topic_title, max_members=1)
    name = message.from_user.full_name
    await add_member(topic_id, message.from_user.id, name)

    await state.clear()
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📋 К каталогу тем", callback_data=f"open_list_{list_id}_0")],
        [InlineKeyboardButton(text="◀️ Главное меню", callback_data="main_menu")]
    ])
    await message.answer(f"✅ Ваша тема <b>«{topic_title}»</b> добавлена в каталог и забронирована за вами!", parse_mode="HTML", reply_markup=kb)

# Карточка темы
@router.callback_query(F.data.startswith("topic_"))
async def cb_topic_details(callback: CallbackQuery):
    parts = callback.data.split("_")
    topic_id = int(parts[1])
    page = int(parts[2]) if len(parts) > 2 else 0

    topic = await get_topic_by_id(topic_id)
    if not topic:
        await callback.answer("Тема не найдена")
        return

    list_id, title, max_m = topic["list_id"], topic["title"], topic["max_members"]
    members = await get_members(topic_id)
    cur_m = len(members)

    text = f"📌 <b>{title}</b>\nЛимит мест: {cur_m}/{max_m}\n\n<b>Записаны:</b>\n"
    for m in members:
        text += f"• {m['user_name']}\n"
    if not members:
        text += "— Место свободно\n"

    user_ids = [m["user_id"] for m in members]
    is_user_in_this_topic = callback.from_user.id in user_ids
    user_topic_in_list = await get_user_topic_in_list(list_id, callback.from_user.id)

    kb = []
    if is_user_in_this_topic:
        kb.append([InlineKeyboardButton(text="❌ Отказаться от темы", callback_data=f"leave_{topic_id}_{page}")])
    elif user_topic_in_list:
        text += f"\n⚠️ <i>Вы уже записаны на тему «{user_topic_in_list['title']}». Можно выбрать только одну тему.</i>"
    elif cur_m < max_m:
        kb.append([InlineKeyboardButton(text="✅ Записаться на тему", callback_data=f"take_{topic_id}_{page}")])

    kb.append([InlineKeyboardButton(text="◀️ Назад к списку", callback_data=f"open_list_{list_id}_{page}")])
    await callback.message.edit_text(text, parse_mode="HTML", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("take_"))
async def cb_take_topic(callback: CallbackQuery):
    parts = callback.data.split("_")
    topic_id = int(parts[1])
    name = callback.from_user.full_name
    success, msg = await add_member(topic_id, callback.from_user.id, name)
    if success:
        await callback.answer("✅ Вы успешно записались!")
    else:
        await callback.answer(msg, show_alert=True)
    await cb_topic_details(callback)

@router.callback_query(F.data.startswith("leave_"))
async def cb_leave_topic(callback: CallbackQuery):
    parts = callback.data.split("_")
    topic_id = int(parts[1])
    await remove_member(topic_id, callback.from_user.id)
    await callback.answer("Вы отказались от темы.")
    await cb_topic_details(callback)

@router.callback_query(F.data == "my_topics")
async def cb_my_topics(callback: CallbackQuery):
    topics = await get_user_topics(callback.from_user.id)
    back_kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Главное меню", callback_data="main_menu")]])
    if not topics:
        await callback.message.edit_text("У вас пока нет выбранных тем.", reply_markup=back_kb)
        return

    text = "<b>Ваши выбранные темы:</b>\n\n"
    kb = []
    for topic in topics:
        tid, t_title, l_title = topic["id"], topic["title"], topic["list_title"]
        text += f"• <b>{l_title}</b>: {t_title}\n"
        kb.append([InlineKeyboardButton(text=f"❌ Отказаться: {t_title[:20]}...", callback_data=f"leave_{tid}_0")])
    kb.append([InlineKeyboardButton(text="◀️ Главное меню", callback_data="main_menu")])

    await callback.message.edit_text(text, parse_mode="HTML", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

# --- ВЫГРУЗКА EXCEL ---
@router.callback_query(F.data == "download_excel")
async def cb_download_excel(callback: CallbackQuery):
    await callback.answer("Формирую файл...")
    wb = Workbook()
    ws = wb.active
    ws.title = "Реестр тем"

    headers = ["Предмет", "Тема", "Лимит мест", "Студент"]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    lists = await get_all_lists()
    for l in lists:
        lid, l_title = l["id"], l["title"]
        topics = await get_topics(lid)
        for t in topics:
            tid, t_title, max_m = t["id"], t["title"], t["max_members"]
            members = await get_members(tid)
            if members:
                for m in members:
                    ws.append([l_title, t_title, max_m, m["user_name"]])
            else:
                ws.append([l_title, t_title, max_m, "— Свободно —"])

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    file = BufferedInputFile(buf.getvalue(), filename="Реестр_тем.xlsx")
    await callback.message.answer_document(document=file, caption="📊 Актуальный реестр тем")

# --- ВЕБ-СЕРВЕР PING ДЛЯ UPTIMEROBOT ---
async def handle_ping(request):
    return web.Response(text="Bot is running!")

async def start_web_server():
    app = web.Application()
    app.router.add_get('/', handle_ping)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, '0.0.0.0', PORT)
    await site.start()

# --- ТОЧКА ВХОДА ---
async def main():
    await init_db()
    await start_web_server()
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
