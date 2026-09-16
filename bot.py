import os
import io
import math
import logging
import asyncio
import psycopg2
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

# --- БАЗА ДАННЫХ ---
def get_db_connection():
    if DATABASE_URL:
        return psycopg2.connect(DATABASE_URL)
    import sqlite3
    return sqlite3.connect("bot.db")

def query_placeholder():
    return "%s" if DATABASE_URL else "?"

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    if DATABASE_URL:
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS lists (
                id SERIAL PRIMARY KEY,
                title TEXT NOT NULL
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS topics (
                id SERIAL PRIMARY KEY,
                list_id INTEGER REFERENCES lists(id) ON DELETE CASCADE,
                title TEXT NOT NULL,
                max_members INTEGER DEFAULT 1
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS topic_members (
                id SERIAL PRIMARY KEY,
                topic_id INTEGER REFERENCES topics(id) ON DELETE CASCADE,
                user_id BIGINT NOT NULL,
                user_name TEXT NOT NULL
            )
        """)
    else:
        cursor.execute("PRAGMA foreign_keys = ON")
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS lists (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS topics (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                list_id INTEGER NOT NULL,
                title TEXT NOT NULL,
                max_members INTEGER DEFAULT 1,
                FOREIGN KEY (list_id) REFERENCES lists (id) ON DELETE CASCADE
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS topic_members (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                topic_id INTEGER NOT NULL,
                user_id INTEGER NOT NULL,
                user_name TEXT NOT NULL,
                FOREIGN KEY (topic_id) REFERENCES topics (id) ON DELETE CASCADE
            )
        """)
    conn.commit()
    cursor.close()
    conn.close()

# --- ФУНКЦИИ ВЫБОРКИ И ЗАПИСИ ---
def get_all_lists():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, title FROM lists ORDER BY id DESC")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return rows

def get_list_by_id(list_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"SELECT id, title FROM lists WHERE id = {p}", (list_id,))
    row = cursor.fetchone()
    cursor.close()
    conn.close()
    return row

def get_topics(list_id, search_query=None):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    if search_query:
        if DATABASE_URL:
            cursor.execute(f"SELECT id, title, max_members FROM topics WHERE list_id = {p} AND title ILIKE {p} ORDER BY id ASC", (list_id, f"%{search_query}%"))
        else:
            cursor.execute(f"SELECT id, title, max_members FROM topics WHERE list_id = {p} AND title LIKE {p} ORDER BY id ASC", (list_id, f"%{search_query}%"))
    else:
        cursor.execute(f"SELECT id, title, max_members FROM topics WHERE list_id = {p} ORDER BY id ASC", (list_id,))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return rows

def get_topic_by_id(topic_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"SELECT id, list_id, title, max_members FROM topics WHERE id = {p}", (topic_id,))
    row = cursor.fetchone()
    cursor.close()
    conn.close()
    return row

def get_members(topic_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"SELECT user_id, user_name FROM topic_members WHERE topic_id = {p}", (topic_id,))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return rows

def get_user_topic_in_list(list_id, user_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"""
        SELECT t.id, t.title 
        FROM topic_members tm
        JOIN topics t ON tm.topic_id = t.id
        WHERE t.list_id = {p} AND tm.user_id = {p}
    """, (list_id, user_id))
    row = cursor.fetchone()
    cursor.close()
    conn.close()
    return row

def add_member(topic_id, user_id, user_name):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()

    cursor.execute(f"SELECT list_id, max_members FROM topics WHERE id = {p}", (topic_id,))
    res = cursor.fetchone()
    if not res:
        cursor.close()
        conn.close()
        return False, "Тема не найдена."
    list_id, max_m = res

    cursor.execute(f"""
        SELECT t.title 
        FROM topic_members tm
        JOIN topics t ON tm.topic_id = t.id
        WHERE t.list_id = {p} AND tm.user_id = {p}
    """, (list_id, user_id))
    existing = cursor.fetchone()
    if existing:
        cursor.close()
        conn.close()
        return False, f"Вы уже записаны на тему «{existing[0]}». Можно выбрать только одну тему!"

    cursor.execute(f"SELECT COUNT(*) FROM topic_members WHERE topic_id = {p}", (topic_id,))
    cur_m = cursor.fetchone()[0]
    if cur_m >= max_m:
        cursor.close()
        conn.close()
        return False, "Места на эту тему уже закончились!"

    cursor.execute(f"INSERT INTO topic_members (topic_id, user_id, user_name) VALUES ({p}, {p}, {p})", (topic_id, user_id, user_name))
    conn.commit()
    cursor.close()
    conn.close()
    return True, "Успешно!"

def remove_member(topic_id, user_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"DELETE FROM topic_members WHERE topic_id = {p} AND user_id = {p}", (topic_id, user_id))
    conn.commit()
    cursor.close()
    conn.close()

def clear_topic_members(topic_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"DELETE FROM topic_members WHERE topic_id = {p}", (topic_id,))
    conn.commit()
    cursor.close()
    conn.close()

def get_user_topics(user_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"""
        SELECT t.id, t.title, l.title 
        FROM topic_members tm
        JOIN topics t ON tm.topic_id = t.id
        JOIN lists l ON t.list_id = l.id
        WHERE tm.user_id = {p}
    """, (user_id,))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return rows

def create_list(title):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    if DATABASE_URL:
        cursor.execute(f"INSERT INTO lists (title) VALUES ({p}) RETURNING id", (title,))
        list_id = cursor.fetchone()[0]
    else:
        cursor.execute(f"INSERT INTO lists (title) VALUES ({p})", (title,))
        list_id = cursor.lastrowid
    conn.commit()
    cursor.close()
    conn.close()
    return list_id

def create_topic(list_id, title, max_members=1):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    if DATABASE_URL:
        cursor.execute(f"INSERT INTO topics (list_id, title, max_members) VALUES ({p}, {p}, {p}) RETURNING id", (list_id, title, max_members))
        topic_id = cursor.fetchone()[0]
    else:
        cursor.execute(f"INSERT INTO topics (list_id, title, max_members) VALUES ({p}, {p}, {p})", (list_id, title, max_members))
        topic_id = cursor.lastrowid
    conn.commit()
    cursor.close()
    conn.close()
    return topic_id

def update_topic_title(topic_id, new_title):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"UPDATE topics SET title = {p} WHERE id = {p}", (new_title, topic_id))
    conn.commit()
    cursor.close()
    conn.close()

def update_topic_max_members(topic_id, new_max):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"UPDATE topics SET max_members = {p} WHERE id = {p}", (new_max, topic_id))
    conn.commit()
    cursor.close()
    conn.close()

def delete_topic(topic_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"DELETE FROM topics WHERE id = {p}", (topic_id,))
    conn.commit()
    cursor.close()
    conn.close()

def batch_create_topics(topics_data):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    query = f"INSERT INTO topics (list_id, title, max_members) VALUES ({p}, {p}, {p})"
    cursor.executemany(query, topics_data)
    conn.commit()
    cursor.close()
    conn.close()

def delete_list(list_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"DELETE FROM lists WHERE id = {p}", (list_id,))
    conn.commit()
    cursor.close()
    conn.close()

# --- МЕНЮ И КЛАВИАТУРЫ ---
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

# Генератор страниц списка для студентов и старосты
def build_topics_page_keyboard(list_id, topics, user_id, page=0, is_search=False, is_admin=False):
    total_topics = len(topics)
    total_pages = max(1, math.ceil(total_topics / PAGE_SIZE))
    page = max(0, min(page, total_pages - 1))

    start_idx = page * PAGE_SIZE
    end_idx = start_idx + PAGE_SIZE
    current_page_topics = topics[start_idx:end_idx]

    kb = []
    for tid, title, max_m in current_page_topics:
        members = get_members(tid)
        cur_m = len(members)
        is_my_topic = any(m[0] == user_id for m in members)

        if is_my_topic:
            status = "⭐️ ВЫ"
        elif cur_m >= max_m:
            status = "🔴 Занято"
        else:
            status = f"🟢 ({cur_m}/{max_m})"

        display_title = title if len(title) <= 35 else title[:32] + "..."
        btn_text = f"{display_title} — {status}"
        kb.append([InlineKeyboardButton(text=btn_text, callback_data=f"topic_{tid}_{page}")])

    # Кнопки страниц
    nav_row = []
    if page > 0:
        nav_row.append(InlineKeyboardButton(text="⬅️ Назад", callback_data=f"page_{list_id}_{page - 1}"))
    if total_pages > 1:
        nav_row.append(InlineKeyboardButton(text=f"Стр. {page + 1}/{total_pages}", callback_data="noop"))
    if page < total_pages - 1:
        nav_row.append(InlineKeyboardButton(text="Вперёд ➡️", callback_data=f"page_{list_id}_{page + 1}"))

    if nav_row:
        kb.append(nav_row)

    # Дополнительные инструменты
    bottom_row = []
    if is_search:
        bottom_row.append(InlineKeyboardButton(text="🔄 Сбросить поиск", callback_data=f"open_list_{list_id}_0"))
    else:
        bottom_row.append(InlineKeyboardButton(text="🔍 Поиск темы", callback_data=f"search_topic_{list_id}"))

    user_chosen = get_user_topic_in_list(list_id, user_id)
    if not user_chosen:
        bottom_row.append(InlineKeyboardButton(text="💡 «Другое»", callback_data=f"custom_topic_{list_id}"))

    if bottom_row:
        kb.append(bottom_row)

    # Если староста просматривает каталог, даем быструю кнопку добавления тем в этот предмет
    if is_admin:
        kb.append([InlineKeyboardButton(text="➕ Добавить темы в этот список", callback_data=f"admin_append_to_{list_id}")])

    kb.append([InlineKeyboardButton(text="◀️ К списку предметов", callback_data="view_lists")])
    return InlineKeyboardMarkup(inline_keyboard=kb)

# Генератор страниц тем для панели старосты
def build_admin_topics_page_keyboard(list_id, topics, page=0):
    total_topics = len(topics)
    total_pages = max(1, math.ceil(total_topics / PAGE_SIZE))
    page = max(0, min(page, total_pages - 1))

    start_idx = page * PAGE_SIZE
    end_idx = start_idx + PAGE_SIZE
    current_page_topics = topics[start_idx:end_idx]

    kb = []
    for tid, title, max_m in current_page_topics:
        members = get_members(tid)
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

# --- БАЗОВЫЕ ХЕНДЛЕРЫ ---
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

# --- МЕНЮ СТАРОСТЫ (АДМИН) ---
@router.callback_query(F.data == "admin_menu")
async def cb_admin_menu(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    if ADMIN_ID != 0 and callback.from_user.id != ADMIN_ID:
        await callback.answer("У вас нет прав старосты.", show_alert=True)
        return
    await callback.message.edit_text("⚙️ <b>Панель старосты:</b>\nВыберите действие:", parse_mode="HTML", reply_markup=get_admin_menu())

# Создание нового предмета
@router.callback_query(F.data == "admin_create_list")
async def cb_admin_create_list(callback: CallbackQuery, state: FSMContext):
    await state.set_state(AdminStates.waiting_for_list_title)
    await callback.message.edit_text(
        "✏️ Введите название предмета (например: <i>Макроэкономика</i>):",
        parse_mode="HTML"
    )

@router.message(AdminStates.waiting_for_list_title)
async def process_list_title(message: Message, state: FSMContext):
    title = message.text.strip()
    list_id = create_list(title)
    await state.update_data(list_id=list_id, list_title=title)
    await state.set_state(AdminStates.waiting_for_topics)
    await message.answer(
        f"Предмет <b>«{title}»</b> создан!\n\n"
        "Отправьте список тем сообщением (каждая тема с новой строки).\n"
        "<i>Если на тему нужно несколько человек, укажите в конце цифру в скобках, например: Тема доклада (2)</i>",
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
        batch_create_topics(topics_to_insert)

    await state.clear()
    is_admin = (ADMIN_ID == 0) or (message.from_user.id == ADMIN_ID)
    await message.answer(
        f"✅ Успешно добавлено {len(topics_to_insert)} тем к предмету <b>«{data.get('list_title')}»</b>!",
        parse_mode="HTML",
        reply_markup=get_main_menu(is_admin)
    )

# Добавление тем в существующий предмет
@router.callback_query(F.data == "admin_append_topics_choose")
async def cb_admin_append_topics_choose(callback: CallbackQuery):
    lists = get_all_lists()
    if not lists:
        await callback.message.edit_text("Нет созданных предметов.", reply_markup=get_admin_menu())
        return
    kb = [[InlineKeyboardButton(text=t, callback_data=f"admin_append_to_{lid}")] for lid, t in lists]
    kb.append([InlineKeyboardButton(text="◀️ Назад", callback_data="admin_menu")])
    await callback.message.edit_text("Выберите предмет, в который нужно добавить темы:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("admin_append_to_"))
async def cb_admin_append_to(callback: CallbackQuery, state: FSMContext):
    list_id = int(callback.data.split("_")[3])
    item = get_list_by_id(list_id)
    await state.update_data(append_list_id=list_id, append_list_title=item[1])
    await state.set_state(AdminStates.waiting_for_append_topics)
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Отмена", callback_data="admin_menu")]])
    await callback.message.edit_text(
        f"📥 Отправьте новые темы для предмета <b>«{item[1]}»</b> (каждая тема с новой строки):\n\n"
        "<i>Можно отправить как одну тему, так и список из нескольких.</i>",
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
        batch_create_topics(topics_to_insert)

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

# Редактирование тем
@router.callback_query(F.data == "admin_edit_topics_choose")
async def cb_admin_edit_topics_choose(callback: CallbackQuery):
    lists = get_all_lists()
    if not lists:
        await callback.message.edit_text("Нет созданных предметов.", reply_markup=get_admin_menu())
        return
    kb = [[InlineKeyboardButton(text=t, callback_data=f"admin_open_edit_list_{lid}_0")] for lid, t in lists]
    kb.append([InlineKeyboardButton(text="◀️ Назад", callback_data="admin_menu")])
    await callback.message.edit_text("Выберите предмет для редактирования тем:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("admin_open_edit_list_"))
async def cb_admin_open_edit_list(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    parts = callback.data.split("_")
    list_id = int(parts[4])
    page = int(parts[5]) if len(parts) > 5 else 0

    topics = get_topics(list_id)
    item = get_list_by_id(list_id)
    if not topics:
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="➕ Добавить темы в этот предмет", callback_data=f"admin_append_to_{list_id}")],
            [InlineKeyboardButton(text="◀️ Назад", callback_data="admin_edit_topics_choose")]
        ])
        await callback.message.edit_text(f"В предмете «{item[1]}» пока нет тем.", reply_markup=kb)
        return

    kb = build_admin_topics_page_keyboard(list_id, topics, page=page)
    await callback.message.edit_text(
        f"✏️ <b>Редактирование тем: «{item[1]}»</b>\nВыберите тему для изменения:",
        parse_mode="HTML",
        reply_markup=kb
    )

@router.callback_query(F.data.startswith("admin_page_"))
async def cb_admin_page(callback: CallbackQuery):
    parts = callback.data.split("_")
    list_id = int(parts[2])
    page = int(parts[3])

    topics = get_topics(list_id)
    item = get_list_by_id(list_id)
    kb = build_admin_topics_page_keyboard(list_id, topics, page=page)
    await callback.message.edit_text(
        f"✏️ <b>Редактирование тем: «{item[1]}»</b>\nВыберите тему для изменения:",
        parse_mode="HTML",
        reply_markup=kb
    )

# Детальное меню редактирования темы
@router.callback_query(F.data.startswith("admin_topic_"))
async def cb_admin_topic_detail(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    parts = callback.data.split("_")
    topic_id = int(parts[2])
    page = int(parts[3]) if len(parts) > 3 else 0

    topic = get_topic_by_id(topic_id)
    if not topic:
        await callback.answer("Тема не найдена")
        return

    _, list_id, title, max_m = topic
    members = get_members(topic_id)

    text = f"⚙️ <b>Управление темой:</b>\n\n<b>Название:</b> {title}\n<b>Лимит мест:</b> {len(members)}/{max_m}\n\n<b>Записаны:</b>\n"
    if members:
        for _, name in members:
            text += f"• {name}\n"
    else:
        text += "— Свободно\n"

    kb = [
        [InlineKeyboardButton(text="✏️ Изменить название", callback_data=f"edit_title_{topic_id}_{page}")],
        [InlineKeyboardButton(text="👥 Изменить лимит мест", callback_data=f"edit_limit_{topic_id}_{page}")],
        [InlineKeyboardButton(text="🔄 Снять всех студентов", callback_data=f"clear_mem_{topic_id}_{page}")],
        [InlineKeyboardButton(text="🗑 Удалить тему полностью", callback_data=f"del_topic_{topic_id}_{page}")],
        [InlineKeyboardButton(text="◀️ Назад к темам", callback_data=f"admin_open_edit_list_{list_id}_{page}")]
    ]
    await callback.message.edit_text(text, parse_mode="HTML", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

# Изменение названия
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

    update_topic_title(topic_id, new_title)
    await state.clear()
    await message.answer("✅ Название темы обновлено!")

    topic = get_topic_by_id(topic_id)
    _, list_id, title, max_m = topic
    members = get_members(topic_id)
    text = f"⚙️ <b>Управление темой:</b>\n\n<b>Название:</b> {title}\n<b>Лимит мест:</b> {len(members)}/{max_m}\n\n<b>Записаны:</b>\n"
    for _, name in members:
        text += f"• {name}\n"
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

# Изменение лимита мест
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

    update_topic_max_members(topic_id, new_max)
    await state.clear()
    await message.answer(f"✅ Лимит мест изменен на {new_max}!")

    topic = get_topic_by_id(topic_id)
    _, list_id, title, max_m = topic
    members = get_members(topic_id)
    text = f"⚙️ <b>Управление темой:</b>\n\n<b>Название:</b> {title}\n<b>Лимит мест:</b> {len(members)}/{max_m}\n\n<b>Записаны:</b>\n"
    for _, name in members:
        text += f"• {name}\n"
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

# Снять всех с темы
@router.callback_query(F.data.startswith("clear_mem_"))
async def cb_clear_mem(callback: CallbackQuery):
    parts = callback.data.split("_")
    topic_id = int(parts[2])
    clear_topic_members(topic_id)
    await callback.answer("Все студенты сняты с темы!")
    await cb_admin_topic_detail(callback, None)

# Удаление отдельной темы
@router.callback_query(F.data.startswith("del_topic_"))
async def cb_del_topic(callback: CallbackQuery):
    parts = callback.data.split("_")
    topic_id = int(parts[2])
    page = int(parts[3])

    topic = get_topic_by_id(topic_id)
    list_id = topic[1]
    delete_topic(topic_id)
    await callback.answer("Тема удалена!")

    topics = get_topics(list_id)
    item = get_list_by_id(list_id)
    kb = build_admin_topics_page_keyboard(list_id, topics, page=page)
    await callback.message.edit_text(
        f"✏️ <b>Редактирование тем: «{item[1]}»</b>\nТема была удалена. Выберите действие:",
        parse_mode="HTML",
        reply_markup=kb
    )

# Удаление целого предмета
@router.callback_query(F.data == "admin_delete_list")
async def cb_admin_delete_list(callback: CallbackQuery):
    lists = get_all_lists()
    if not lists:
        await callback.message.edit_text("Нет доступных предметов для удаления.", reply_markup=get_admin_menu())
        return
    kb = [[InlineKeyboardButton(text=f"❌ {t}", callback_data=f"del_list_{lid}")] for lid, t in lists]
    kb.append([InlineKeyboardButton(text="◀️ Назад", callback_data="admin_menu")])
    await callback.message.edit_text("Выберите предмет, который хотите полностью удалить:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("del_list_"))
async def cb_confirm_del_list(callback: CallbackQuery):
    list_id = int(callback.data.split("_")[2])
    delete_list(list_id)
    await callback.answer("Предмет и все темы удалены!")
    await cb_admin_delete_list(callback)

# --- СТУДЕНЧЕСКИЙ КАТАЛОГ И ЗАПИСЬ ---
@router.callback_query(F.data == "view_lists")
async def cb_view_lists(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    lists = get_all_lists()
    if not lists:
        back_kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Назад", callback_data="main_menu")]])
        await callback.message.edit_text("Пока нет доступных предметов.", reply_markup=back_kb)
        return
    kb = [[InlineKeyboardButton(text=t, callback_data=f"open_list_{lid}_0")] for lid, t in lists]
    kb.append([InlineKeyboardButton(text="◀️ Главное меню", callback_data="main_menu")])
    await callback.message.edit_text("Выберите предмет:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("open_list_"))
async def cb_open_list(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    parts = callback.data.split("_")
    list_id = int(parts[2])
    page = int(parts[3]) if len(parts) > 3 else 0

    topics = get_topics(list_id)
    is_admin = (ADMIN_ID == 0) or (callback.from_user.id == ADMIN_ID)

    if not topics:
        kb_rows = [
            [InlineKeyboardButton(text="💡 Предложить свою тему («Другое»)", callback_data=f"custom_topic_{list_id}")],
        ]
        if is_admin:
            kb_rows.append([InlineKeyboardButton(text="➕ Добавить темы в этот список", callback_data=f"admin_append_to_{list_id}")])
        kb_rows.append([InlineKeyboardButton(text="◀️ Назад", callback_data="view_lists")])
        
        await callback.message.edit_text("В этом предмете пока нет тем. Вы можете предложить тему или староста может добавить список:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb_rows))
        return

    user_chosen = get_user_topic_in_list(list_id, callback.from_user.id)
    header_text = f"📋 <b>Каталог тем</b> (Всего тем: {len(topics)})\n"
    if user_chosen:
        header_text += f"\n⚠️ <i>Вы уже записаны на тему: «{user_chosen[1]}»</i>"

    kb = build_topics_page_keyboard(list_id, topics, callback.from_user.id, page=page, is_admin=is_admin)
    await callback.message.edit_text(header_text, parse_mode="HTML", reply_markup=kb)

@router.callback_query(F.data.startswith("page_"))
async def cb_change_page(callback: CallbackQuery):
    parts = callback.data.split("_")
    list_id = int(parts[1])
    page = int(parts[2])

    topics = get_topics(list_id)
    is_admin = (ADMIN_ID == 0) or (callback.from_user.id == ADMIN_ID)
    user_chosen = get_user_topic_in_list(list_id, callback.from_user.id)
    header_text = f"📋 <b>Каталог тем</b> (Всего тем: {len(topics)})\n"
    if user_chosen:
        header_text += f"\n⚠️ <i>Вы уже записаны на тему: «{user_chosen[1]}»</i>"

    kb = build_topics_page_keyboard(list_id, topics, callback.from_user.id, page=page, is_admin=is_admin)
    await callback.message.edit_text(header_text, parse_mode="HTML", reply_markup=kb)

# Поиск темы
@router.callback_query(F.data.startswith("search_topic_"))
async def cb_search_topic(callback: CallbackQuery, state: FSMContext):
    list_id = int(callback.data.split("_")[2])
    await state.update_data(search_list_id=list_id)
    await state.set_state(StudentStates.waiting_for_search_query)
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Отмена", callback_data=f"open_list_{list_id}_0")]])
    await callback.message.edit_text(
        "🔍 <b>Поиск по темам</b>\n\nОтправьте ключевое слово или часть названия темы:",
        parse_mode="HTML",
        reply_markup=kb
    )

@router.message(StudentStates.waiting_for_search_query)
async def process_search_query(message: Message, state: FSMContext):
    query = message.text.strip()
    data = await state.get_data()
    list_id = data.get("search_list_id")
    await state.clear()

    topics = get_topics(list_id, search_query=query)
    if not topics:
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🔄 Сбросить поиск", callback_data=f"open_list_{list_id}_0")],
            [InlineKeyboardButton(text="🔍 Искать снова", callback_data=f"search_topic_{list_id}")]
        ])
        await message.answer(f"По запросу «{query}» ничего не найдено.", reply_markup=kb)
        return

    is_admin = (ADMIN_ID == 0) or (message.from_user.id == ADMIN_ID)
    kb = build_topics_page_keyboard(list_id, topics, message.from_user.id, page=0, is_search=True, is_admin=is_admin)
    await message.answer(
        f"🔍 Результаты поиска по запросу «<b>{query}</b>» (найдено: {len(topics)}):",
        parse_mode="HTML",
        reply_markup=kb
    )

# Добавление темы студентом («Другое»)
@router.callback_query(F.data.startswith("custom_topic_"))
async def cb_custom_topic_start(callback: CallbackQuery, state: FSMContext):
    list_id = int(callback.data.split("_")[2])
    user_chosen = get_user_topic_in_list(list_id, callback.from_user.id)
    if user_chosen:
        await callback.answer("Вы уже записаны на тему в этом предмете!", show_alert=True)
        return

    await state.update_data(custom_list_id=list_id)
    await state.set_state(StudentStates.waiting_for_custom_topic)
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Отмена", callback_data=f"open_list_{list_id}_0")]])
    await callback.message.edit_text(
        "💡 <b>Предложить свою тему («Другое»)</b>\n\n"
        "Напишите точное название вашей темы сообщением в чат.\n"
        "Она сразу появится в общем списке предмета и автоматически закрепится за вами!",
        parse_mode="HTML",
        reply_markup=kb
    )

@router.message(StudentStates.waiting_for_custom_topic)
async def process_custom_topic_text(message: Message, state: FSMContext):
    topic_title = message.text.strip()
    if not topic_title:
        await message.answer("Пожалуйста, отправьте непустой текст темы.")
        return

    data = await state.get_data()
    list_id = data.get("custom_list_id")

    user_chosen = get_user_topic_in_list(list_id, message.from_user.id)
    if user_chosen:
        await state.clear()
        await message.answer("Вы уже записаны на тему в этом предмете.")
        return

    topic_id = create_topic(list_id, topic_title, max_members=1)
    name = message.from_user.full_name
    add_member(topic_id, message.from_user.id, name)

    await state.clear()
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📋 К каталогу тем", callback_data=f"open_list_{list_id}_0")],
        [InlineKeyboardButton(text="◀️ Главное меню", callback_data="main_menu")]
    ])
    await message.answer(
        f"✅ Ваша тема <b>«{topic_title}»</b> добавлена в каталог и забронирована за вами!",
        parse_mode="HTML",
        reply_markup=kb
    )

# Карточка темы для студента
@router.callback_query(F.data.startswith("topic_"))
async def cb_topic_details(callback: CallbackQuery):
    parts = callback.data.split("_")
    topic_id = int(parts[1])
    page = int(parts[2]) if len(parts) > 2 else 0

    topic = get_topic_by_id(topic_id)
    if not topic:
        await callback.answer("Тема не найдена")
        return

    _, list_id, title, max_m = topic
    members = get_members(topic_id)
    cur_m = len(members)

    text = f"📌 <b>{title}</b>\nЛимит мест: {cur_m}/{max_m}\n\n<b>Записаны:</b>\n"
    if members:
        for _, name in members:
            text += f"• {name}\n"
    else:
        text += "— Место свободно\n"

    user_ids = [m[0] for m in members]
    is_user_in_this_topic = callback.from_user.id in user_ids
    user_topic_in_list = get_user_topic_in_list(list_id, callback.from_user.id)

    kb = []
    if is_user_in_this_topic:
        kb.append([InlineKeyboardButton(text="❌ Отказаться от темы", callback_data=f"leave_{topic_id}_{page}")])
    elif user_topic_in_list:
        text += f"\n⚠️ <i>Вы уже записаны на тему «{user_topic_in_list[1]}». Можно выбрать только одну тему.</i>"
    elif cur_m < max_m:
        kb.append([InlineKeyboardButton(text="✅ Записаться на тему", callback_data=f"take_{topic_id}_{page}")])

    kb.append([InlineKeyboardButton(text="◀️ Назад к списку", callback_data=f"open_list_{list_id}_{page}")])
    await callback.message.edit_text(text, parse_mode="HTML", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("take_"))
async def cb_take_topic(callback: CallbackQuery):
    parts = callback.data.split("_")
    topic_id = int(parts[1])
    name = callback.from_user.full_name
    success, msg = add_member(topic_id, callback.from_user.id, name)
    if success:
        await callback.answer("✅ Вы успешно записались!")
    else:
        await callback.answer(msg, show_alert=True)
    await cb_topic_details(callback)

@router.callback_query(F.data.startswith("leave_"))
async def cb_leave_topic(callback: CallbackQuery):
    parts = callback.data.split("_")
    topic_id = int(parts[1])
    remove_member(topic_id, callback.from_user.id)
    await callback.answer("Вы отказались от темы.")
    await cb_topic_details(callback)

@router.callback_query(F.data == "my_topics")
async def cb_my_topics(callback: CallbackQuery):
    topics = get_user_topics(callback.from_user.id)
    back_kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Главное меню", callback_data="main_menu")]])
    if not topics:
        await callback.message.edit_text("У вас пока нет выбранных тем.", reply_markup=back_kb)
        return

    text = "<b>Ваши выбранные темы:</b>\n\n"
    kb = []
    for tid, t_title, l_title in topics:
        text += f"• <b>{l_title}</b>: {t_title}\n"
        kb.append([InlineKeyboardButton(text=f"❌ Отказаться: {t_title[:20]}...", callback_data=f"leave_{tid}_0")])
    kb.append([InlineKeyboardButton(text="◀️ Главное меню", callback_data="main_menu")])

    await callback.message.edit_text(text, parse_mode="HTML", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

# --- ВЫГРУЗКА EXCEL ---
@router.callback_query(F.data == "download_excel")
async def cb_download_excel(callback: CallbackQuery):
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

    lists = get_all_lists()
    for lid, l_title in lists:
        topics = get_topics(lid)
        for tid, t_title, max_m in topics:
            members = get_members(tid)
            if members:
                for _, name in members:
                    ws.append([l_title, t_title, max_m, name])
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
    await callback.answer()

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
    init_db()
    await start_web_server()
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
