import os
import io
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

logging.basicConfig(level=logging.INFO)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
router = Router()
dp.include_router(router)

# --- FSM СОСТОЯНИЯ ---
class AdminStates(StatesGroup):
    waiting_for_list_title = State()
    waiting_for_topics = State()

# --- РАБОТА С БАЗОЙ ДАННЫХ ---
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

# --- ФУНКЦИИ ДАННЫХ ---
def get_all_lists():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, title FROM lists ORDER BY id DESC")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return rows

def get_topics(list_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"SELECT id, title, max_members FROM topics WHERE list_id = {p} ORDER BY id ASC", (list_id,))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return rows

def get_members(topic_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"SELECT user_id, user_name FROM topic_members WHERE topic_id = {p}", (topic_id,))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return rows

def add_member(topic_id, user_id, user_name):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"SELECT max_members FROM topics WHERE id = {p}", (topic_id,))
    res = cursor.fetchone()
    if not res:
        cursor.close()
        conn.close()
        return False
    max_m = res[0]

    cursor.execute(f"SELECT COUNT(*) FROM topic_members WHERE topic_id = {p}", (topic_id,))
    cur_m = cursor.fetchone()[0]
    if cur_m >= max_m:
        cursor.close()
        conn.close()
        return False

    cursor.execute(f"INSERT INTO topic_members (topic_id, user_id, user_name) VALUES ({p}, {p}, {p})", (topic_id, user_id, user_name))
    conn.commit()
    cursor.close()
    conn.close()
    return True

def remove_member(topic_id, user_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"DELETE FROM topic_members WHERE topic_id = {p} AND user_id = {p}", (topic_id, user_id))
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
    cursor.execute(f"INSERT INTO topics (list_id, title, max_members) VALUES ({p}, {p}, {p})", (list_id, title, max_members))
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
        [InlineKeyboardButton(text="➕ Создать список тем", callback_data="admin_create_list")],
        [InlineKeyboardButton(text="🗑 Удалить список тем", callback_data="admin_delete_list")],
        [InlineKeyboardButton(text="◀️ Главное меню", callback_data="main_menu")]
    ])

# --- ОБЩИЕ ХЕНДЛЕРЫ ---
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

# --- ХЕНДЛЕРЫ МЕНЮ СТАРОСТЫ ---
@router.callback_query(F.data == "admin_menu")
async def cb_admin_menu(callback: CallbackQuery):
    if ADMIN_ID != 0 and callback.from_user.id != ADMIN_ID:
        await callback.answer("У вас нет прав старосты.", show_alert=True)
        return
    await callback.message.edit_text("⚙️ <b>Панель старосты:</b>\nВыберите действие:", parse_mode="HTML", reply_markup=get_admin_menu())

@router.callback_query(F.data == "admin_create_list")
async def cb_admin_create_list(callback: CallbackQuery, state: FSMContext):
    await state.set_state(AdminStates.waiting_for_list_title)
    await callback.message.edit_text(
        "✏️ Введите название предмета или реестра (например: <i>Экономическая теория</i>):",
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
        "Теперь отправьте список тем **одним сообщением**, где каждая тема с новой строки.\n"
        "<i>(Если на тему нужно больше 1 человека, укажите в конце темы цифру в скобках, например: Тема доклада (2))</i>",
        parse_mode="HTML"
    )

@router.message(AdminStates.waiting_for_topics)
async def process_topics(message: Message, state: FSMContext):
    data = await state.get_data()
    list_id = data.get("list_id")
    lines = [line.strip() for line in message.text.split("\n") if line.strip()]

    count = 0
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
        create_topic(list_id, title, max_m)
        count += 1

    await state.clear()
    is_admin = (ADMIN_ID == 0) or (message.from_user.id == ADMIN_ID)
    await message.answer(
        f"✅ Успешно добавлено {count} тем к предмету <b>«{data.get('list_title')}»</b>!",
        parse_mode="HTML",
        reply_markup=get_main_menu(is_admin)
    )

@router.callback_query(F.data == "admin_delete_list")
async def cb_admin_delete_list(callback: CallbackQuery):
    lists = get_all_lists()
    if not lists:
        await callback.message.edit_text("Нет доступных списков для удаления.", reply_markup=get_admin_menu())
        return
    kb = [[InlineKeyboardButton(text=f"❌ {t}", callback_data=f"del_list_{lid}")] for lid, t in lists]
    kb.append([InlineKeyboardButton(text="◀️ Назад", callback_data="admin_menu")])
    await callback.message.edit_text("Выберите предмет, который хотите полностью удалить:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("del_list_"))
async def cb_confirm_del_list(callback: CallbackQuery):
    list_id = int(callback.data.split("_")[2])
    delete_list(list_id)
    await callback.answer("Предмет и темы удалены!")
    await cb_admin_delete_list(callback)

# --- ПРОСМОТР И ЗАПИСЬ НА ТЕМЫ ---
@router.callback_query(F.data == "view_lists")
async def cb_view_lists(callback: CallbackQuery):
    lists = get_all_lists()
    if not lists:
        await callback.message.edit_text(
            "Пока нет доступных предметов.\nСтароста может создать их в меню старосты.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Назад", callback_data="main_menu")]])
        )
        return
    kb = [[InlineKeyboardButton(text=t, callback_data=f"open_list_{lid}")] for lid, t in lists]
    kb.append([InlineKeyboardButton(text="◀️ Главное меню", callback_data="main_menu")])
    await callback.message.edit_text("Выберите предмет:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("open_list_"))
async def cb_open_list(callback: CallbackQuery):
    list_id = int(callback.data.split("_")[2])
    topics = get_topics(list_id)
    if not topics:
        await callback.message.edit_text("В этом списке пока нет тем.", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Назад", callback_data="view_lists")]]))
        return
    
    kb = []
    for tid, title, max_m in topics:
        members = get_members(tid)
        cur_m = len(members)
        status = "🔴" if cur_m >= max_m else f"🟢 ({cur_m}/{max_m})"
        btn_text = f"{title} — {status}"
        kb.append([InlineKeyboardButton(text=btn_text, callback_data=f"topic_{tid}")])
    kb.append([InlineKeyboardButton(text="◀️ К предметам", callback_data="view_lists")])
    await callback.message.edit_text("Выберите тему для записи:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("topic_"))
async def cb_topic_details(callback: CallbackQuery):
    topic_id = int(callback.data.split("_")[1])
    conn = get_db_connection()
    cursor = conn.cursor()
    p = query_placeholder()
    cursor.execute(f"SELECT list_id, title, max_members FROM topics WHERE id = {p}", (topic_id,))
    topic = cursor.fetchone()
    cursor.close()
    conn.close()

    if not topic:
        await callback.answer("Тема не найдена")
        return

    list_id, title, max_m = topic
    members = get_members(topic_id)
    cur_m = len(members)

    text = f"📌 <b>{title}</b>\nЛимит мест: {cur_m}/{max_m}\n\n<b>Записаны:</b>\n"
    if members:
        for _, name in members:
            text += f"• {name}\n"
    else:
        text += "— Место свободно\n"

    user_ids = [m[0] for m in members]
    kb = []
    if callback.from_user.id in user_ids:
        kb.append([InlineKeyboardButton(text="❌ Отказаться от темы", callback_data=f"leave_{topic_id}")])
    elif cur_m < max_m:
        kb.append([InlineKeyboardButton(text="✅ Записаться на тему", callback_data=f"take_{topic_id}")])

    kb.append([InlineKeyboardButton(text="◀️ Назад к списку", callback_data=f"open_list_{list_id}")])
    await callback.message.edit_text(text, parse_mode="HTML", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))

@router.callback_query(F.data.startswith("take_"))
async def cb_take_topic(callback: CallbackQuery):
    topic_id = int(callback.data.split("_")[1])
    name = callback.from_user.full_name
    if add_member(topic_id, callback.from_user.id, name):
        await callback.answer("Вы успешно записались!")
    else:
        await callback.answer("Не удалось записаться (места заняты).", show_alert=True)
    await cb_topic_details(callback)

@router.callback_query(F.data.startswith("leave_"))
async def cb_leave_topic(callback: CallbackQuery):
    topic_id = int(callback.data.split("_")[1])
    remove_member(topic_id, callback.from_user.id)
    await callback.answer("Вы отказались от темы.")
    await cb_topic_details(callback)

@router.callback_query(F.data == "my_topics")
async def cb_my_topics(callback: CallbackQuery):
    topics = get_user_topics(callback.from_user.id)
    if not topics:
        await callback.message.edit_text("У вас пока нет выбранных тем.", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="◀️ Главное меню", callback_data="main_menu")]]))
        return
    
    text = "<b>Ваши выбранные темы:</b>\n\n"
    for _, t_title, l_title in topics:
        text += f"• <b>{l_title}</b>: {t_title}\n"

    await callback.message.edit_text(text, parse_mode="HTML", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb := [[InlineKeyboardButton(text="◀️ Главное меню", callback_data="main_menu")]]))

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
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    file = BufferedInputFile(buf.getvalue(), filename="Реестр_тем.xlsx")
    await callback.message.answer_document(document=file, caption="📊 Актуальный реестр тем")
    await callback.answer()

# --- ВЕБ-СЕРВЕР ДЛЯ RENDER PING ---
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
