import asyncio
import logging
import sqlite3
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import CommandStart, Command
from aiogram.types import (
    Message,
    CallbackQuery,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    FSInputFile
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage

BOT_TOKEN = "8932435756:AAG2EdxoyHyjAQ1cKjn_E3FC4wB6JMipK2g"
ADMIN_ID = 2042219198
DB_NAME = "topics_registry.db"

logging.basicConfig(level=logging.INFO)
bot = Bot(token=BOT_TOKEN)
router = Router()
dp = Dispatcher(storage=MemoryStorage())
dp.include_router(router)

class AdminStates(StatesGroup):
    waiting_for_list_title = State()
    waiting_for_topics = State()
    waiting_for_custom_limit = State()

def init_db():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
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

def get_all_lists():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, title FROM lists ORDER BY id DESC")
        return cursor.fetchall()

def get_list_by_id(list_id: int):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, title FROM lists WHERE id = ?", (list_id,))
        return cursor.fetchone()

def create_list(title: str):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO lists (title) VALUES (?)", (title,))
        conn.commit()
        return cursor.lastrowid

def delete_list(list_id: int):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM topic_members WHERE topic_id IN (SELECT id FROM topics WHERE list_id = ?)", (list_id,))
        cursor.execute("DELETE FROM topics WHERE list_id = ?", (list_id,))
        cursor.execute("DELETE FROM lists WHERE id = ?", (list_id,))
        conn.commit()

def get_free_topics(list_id: int):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT t.id, t.title, t.max_members, COUNT(m.id) as current_cnt
            FROM topics t
            LEFT JOIN topic_members m ON t.id = m.topic_id
            WHERE t.list_id = ?
            GROUP BY t.id
            HAVING current_cnt < t.max_members
        """, (list_id,))
        return cursor.fetchall()

def get_user_topic_in_list(list_id: int, user_id: int):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT t.id, t.title, t.max_members
            FROM topics t
            JOIN topic_members m ON t.id = m.topic_id
            WHERE t.list_id = ? AND m.user_id = ?
        """, (list_id, user_id))
        return cursor.fetchone()

def get_all_user_topics(user_id: int):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT l.title, t.title, t.id, l.id
            FROM topic_members m
            JOIN topics t ON m.topic_id = t.id
            JOIN lists l ON t.list_id = l.id
            WHERE m.user_id = ?
        """, (user_id,))
        return cursor.fetchall()

def get_topic_members(topic_id: int):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT user_name FROM topic_members WHERE topic_id = ?", (topic_id,))
        return [row[0] for row in cursor.fetchall()]

def get_all_topics_report(list_id: int):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, title, max_members FROM topics WHERE list_id = ?", (list_id,))
        topics = cursor.fetchall()
        result = []
        for t_id, title, max_m in topics:
            members = get_topic_members(t_id)
            result.append((t_id, title, max_m, members))
        return result

def update_topic_limit(topic_id: int, new_limit: int):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("UPDATE topics SET max_members = ? WHERE id = ?", (new_limit, topic_id))
        conn.commit()

def join_topic(topic_id: int, user_id: int, user_name: str) -> bool:
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT max_members FROM topics WHERE id = ?", (topic_id,))
        row = cursor.fetchone()
        if not row:
            return False
        max_members = row[0]
        cursor.execute("SELECT COUNT(*) FROM topic_members WHERE topic_id = ?", (topic_id,))
        cnt = cursor.fetchone()[0]
        if cnt >= max_members:
            return False
        cursor.execute(
            "INSERT INTO topic_members (topic_id, user_id, user_name) VALUES (?, ?, ?)",
            (topic_id, user_id, user_name)
        )
        conn.commit()
        return True

def leave_user_topic(list_id: int, user_id: int):
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            DELETE FROM topic_members
            WHERE user_id = ? AND topic_id IN (SELECT id FROM topics WHERE list_id = ?)
        """, (user_id, list_id))
        conn.commit()

def generate_excel_file(list_id: int, list_title: str):
    data = get_all_topics_report(list_id)
    wb = Workbook()
    ws = wb.active
    ws.title = "Реестр"

    headers = ["№", f"Тема: {list_title}", "Состав команды", "Статус"]
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for col_num in range(1, 5):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, (t_id, title, max_m, members) in enumerate(data, start=2):
        status = f"{len(members)} из {max_m}"
        student_names = ", ".join(members) if members else "—"
        ws.append([row_idx - 1, title, student_names, status])

        for col_idx in range(1, 5):
            c = ws.cell(row=row_idx, column=col_idx)
            c.border = thin_border
            if col_idx in [1, 4]:
                c.alignment = Alignment(horizontal="center")

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 20

    filename = f"Реестр_{list_title[:20].replace(' ', '_')}.xlsx"
    wb.save(filename)
    return filename

def get_main_menu(user_id: int):
    buttons = [
        [InlineKeyboardButton(text="📚 Выбрать предмет и тему", callback_data="select_subject_flow")],
        [InlineKeyboardButton(text="👤 Мои выбранные темы", callback_data="my_topics_all")],
        [InlineKeyboardButton(text="📥 Скачать реестр (Excel)", callback_data="export_excel_select")]
    ]
    if user_id == ADMIN_ID:
        buttons.append([InlineKeyboardButton(text="⚙️ Меню старосты", callback_data="admin_panel")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)

@router.message(Command("post_menu"))
async def cmd_post_group_menu(message: Message):
    if message.from_user.id != ADMIN_ID:
        return

    bot_info = await bot.get_me()
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✨ Открыть каталог предметов и тем", url=f"https://t.me/{bot_info.username}?start=catalog")]
    ])
    await message.reply(
        "📢 **Реестр выбора тем по всем предметам**\n\n"
        "Нажмите кнопку ниже, чтобы открыть бота и выбрать свободные темы:",
        reply_markup=kb
    )

@router.message(Command("start", "reestr", "topics"))
async def cmd_start(message: Message):
    user_name = message.from_user.full_name
    
    if message.chat.type in ["group", "supergroup"]:
        bot_info = await bot.get_me()
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="👤 Открыть каталог в боте", url=f"https://t.me/{bot_info.username}?start=catalog")]
        ])
        await message.reply("Для выбора тем перейдите в диалог с ботом:", reply_markup=kb)
        return

    lists = get_all_lists()
    sub_count = len(lists)
    await message.answer(
        f"Привет, {user_name}!\n\n"
        f"Открыто предметов для выбора тем: **{sub_count}**.\n"
        "Вы можете выбрать по одной теме в каждом доступном предмете.",
        reply_markup=get_main_menu(message.from_user.id)
    )

@router.callback_query(F.data == "back_main")
async def back_to_main(call: CallbackQuery):
    await call.message.edit_text("Главное меню каталога тем:", reply_markup=get_main_menu(call.from_user.id))
    await call.answer()

@router.callback_query(F.data == "select_subject_flow")
async def select_subject_flow(call: CallbackQuery):
    lists = get_all_lists()
    if not lists:
        await call.message.edit_text(
            "Сейчас нет доступных предметов.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="🔙 Назад", callback_data="back_main")]
            ])
        )
        return

    buttons = []
    for l_id, title in lists:
        user_topic = get_user_topic_in_list(l_id, call.from_user.id)
        status_mark = "✅ " if user_topic else "📖 "
        buttons.append([InlineKeyboardButton(text=f"{status_mark}{title}", callback_data=f"open_subj_{l_id}")])
        
    buttons.append([InlineKeyboardButton(text="🔙 Назад в меню", callback_data="back_main")])
    await call.message.edit_text("Выберите предмет из списка:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await call.answer()

@router.callback_query(F.data.startswith("open_subj_"))
async def open_subj_handler(call: CallbackQuery):
    list_id = int(call.data.split("_")[2])
    subj = get_list_by_id(list_id)
    if not subj:
        await call.answer("Предмет не найден.", show_alert=True)
        return

    user_topic = get_user_topic_in_list(list_id, call.from_user.id)
    if user_topic:
        t_id, t_title, max_m = user_topic
        members = get_topic_members(t_id)
        team_str = ", ".join(members)
        
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="❌ Отказаться от этой темы", callback_data=f"leave_subj_{list_id}")],
            [InlineKeyboardButton(text="🔙 К списку предметов", callback_data="select_subject_flow")]
        ])
        await call.message.edit_text(
            f"Предмет: **«{subj[1]}»**\n\n"
            f"✅ Вы уже выбрали тему:\n👉 **{t_title}**\n"
            f"👥 Команда ({len(members)}/{max_m}): {team_str}",
            reply_markup=kb
        )
        return

    free = get_free_topics(list_id)
    if not free:
        await call.message.edit_text(
            f"По предмету **«{subj[1]}»** свободных мест больше нет!",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="🔙 К списку предметов", callback_data="select_subject_flow")]
            ])
        )
        return

    buttons = []
    for topic_id, title, max_m, cur_cnt in free:
        slots = f"[{cur_cnt}/{max_m}] " if max_m > 1 else ""
        display_title = title if len(title) <= 35 else title[:32] + "..."
        btn_text = f"📌 {slots}{display_title}"
        buttons.append([InlineKeyboardButton(text=btn_text, callback_data=f"take_{list_id}_{topic_id}")])
        
    buttons.append([InlineKeyboardButton(text="🔙 К списку предметов", callback_data="select_subject_flow")])
    await call.message.edit_text(
        f"Свободные темы по предмету **«{subj[1]}»**:\nНажмите на тему, чтобы записаться:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
    )
    await call.answer()

@router.callback_query(F.data.startswith("take_"))
async def take_topic_handler(call: CallbackQuery):
    parts = call.data.split("_")
    list_id = int(parts[1])
    topic_id = int(parts[2])

    subj = get_list_by_id(list_id)
    if not subj:
        await call.answer("Предмет удален.", show_alert=True)
        return

    user_id = call.from_user.id
    user_name = call.from_user.full_name
    if call.from_user.username:
        user_name += f" (@{call.from_user.username})"

    if get_user_topic_in_list(list_id, user_id):
        await call.answer("Вы уже заняли тему по этому предмету!", show_alert=True)
        await open_subj_handler(call)
        return

    success = join_topic(topic_id, user_id, user_name)
    if success:
        await call.answer("✅ Вы успешно записались!", show_alert=False)
        my_t = get_user_topic_in_list(list_id, user_id)
        members = get_topic_members(topic_id)
        team_str = ", ".join(members)
        
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="📚 Выбрать тему в другом предмете", callback_data="select_subject_flow")],
            [InlineKeyboardButton(text="🔙 В главное меню", callback_data="back_main")]
        ])
        await call.message.edit_text(
            f"✅ **Тема закреплена!**\n\n"
            f"Предмет: **«{subj[1]}»**\n"
            f"📌 **{my_t[1]}**\n"
            f"👥 Состав команды ({len(members)}/{my_t[2]}): {team_str}",
            reply_markup=kb
        )
    else:
        await call.answer("⚠️ Места в этой теме только что закончились!", show_alert=True)
        await open_subj_handler(call)

@router.callback_query(F.data.startswith("leave_subj_"))
async def leave_subject_topic(call: CallbackQuery):
    list_id = int(call.data.split("_")[2])
    leave_user_topic(list_id, call.from_user.id)
    await call.answer("Вы отказались от темы!", show_alert=True)
    await open_subj_handler(call)

@router.callback_query(F.data == "my_topics_all")
async def my_topics_all_handler(call: CallbackQuery):
    user_topics = get_all_user_topics(call.from_user.id)
    if not user_topics:
        await call.message.edit_text(
            "Вы пока не выбрали ни одной темы.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="📚 Выбрать тему", callback_data="select_subject_flow")],
                [InlineKeyboardButton(text="🔙 Назад", callback_data="back_main")]
            ])
        )
        return

    text = "👤 **Ваши выбранные темы:**\n\n"
    buttons = []
    for l_title, t_title, t_id, l_id in user_topics:
        text += f"📖 **{l_title}**\n👉 {t_title}\n\n"
        buttons.append([InlineKeyboardButton(text=f"❌ Отказаться от «{l_title}»", callback_data=f"leave_subj_{l_id}")])
        
    buttons.append([InlineKeyboardButton(text="🔙 Назад в меню", callback_data="back_main")])
    await call.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await call.answer()

@router.callback_query(F.data == "export_excel_select")
async def export_excel_select_handler(call: CallbackQuery):
    lists = get_all_lists()
    if not lists:
        await call.answer("Нет доступных предметов!", show_alert=True)
        return

    buttons = []
    for l_id, title in lists:
        buttons.append([InlineKeyboardButton(text=f"📊 {title}", callback_data=f"dl_excel_{l_id}")])
    buttons.append([InlineKeyboardButton(text="🔙 Назад", callback_data="back_main")])

    await call.message.edit_text("Выберите предмет для скачивания реестра в Excel:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await call.answer()

@router.callback_query(F.data.startswith("dl_excel_"))
async def dl_excel_handler(call: CallbackQuery):
    list_id = int(call.data.split("_")[2])
    subj = get_list_by_id(list_id)
    if not subj:
        await call.answer("Предмет не найден.", show_alert=True)
        return

    topics = get_all_topics_report(list_id)
    if not topics:
        await call.answer("В этом предмете пока нет тем!", show_alert=True)
        return

    file_path = generate_excel_file(list_id, subj[1])
    excel_doc = FSInputFile(file_path)

    await bot.send_document(
        chat_id=call.message.chat.id,
        document=excel_doc,
        caption=f"📊 **Реестр тем: «{subj[1]}»**"
    )
    if os.path.exists(file_path):
        os.remove(file_path)
    await call.answer()

@router.callback_query(F.data == "admin_panel")
async def admin_panel(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        await call.answer("Доступ только для старосты.", show_alert=True)
        return

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📁 Создать новый предмет", callback_data="admin_create_list")],
        [InlineKeyboardButton(text="➕ Добавить темы в предмет", callback_data="admin_add_topics_select")],
        [InlineKeyboardButton(text="👥 Настроить лимит мест у тем", callback_data="admin_limits_select")],
        [InlineKeyboardButton(text="📊 Посмотреть реестр текстом", callback_data="admin_view_select")],
        [InlineKeyboardButton(text="🗑 Удалить предмет", callback_data="admin_delete_select")],
        [InlineKeyboardButton(text="🔙 Назад в меню", callback_data="back_main")]
    ])
    await call.message.edit_text("⚙️ **Панель старосты**:", reply_markup=kb)
    await call.answer()

@router.callback_query(F.data == "admin_create_list")
async def admin_create_list_start(call: CallbackQuery, state: FSMContext):
    if call.from_user.id != ADMIN_ID:
        return
    await call.message.edit_text("Введите название нового предмета или проекта:")
    await state.set_state(AdminStates.waiting_for_list_title)
    await call.answer()

@router.message(AdminStates.waiting_for_list_title)
async def admin_create_list_finish(message: Message, state: FSMContext):
    if message.from_user.id != ADMIN_ID:
        return
    title = message.text.strip()
    create_list(title)
    await state.clear()
    
    await message.reply(
        f"✅ Предмет **«{title}»** создан и доступен студентам в каталоге!\n"
        "Теперь добавьте в него темы через Меню старосты.",
        reply_markup=get_main_menu(message.from_user.id)
    )

@router.callback_query(F.data == "admin_add_topics_select")
async def admin_add_topics_select(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        return
    lists = get_all_lists()
    if not lists:
        await call.answer("Сначала создайте предмет!", show_alert=True)
        return
    buttons = [[InlineKeyboardButton(text=title, callback_data=f"adm_addto_{l_id}")] for l_id, title in lists]
    buttons.append([InlineKeyboardButton(text="🔙 Назад", callback_data="admin_panel")])
    await call.message.edit_text("В какой предмет добавить темы?", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await call.answer()

@router.callback_query(F.data.startswith("adm_addto_"))
async def adm_addto_start(call: CallbackQuery, state: FSMContext):
    if call.from_user.id != ADMIN_ID:
        return
    list_id = int(call.data.split("_")[2])
    await state.update_data(target_list_id=list_id)
    
    subj = get_list_by_id(list_id)
    await call.message.edit_text(
        f"Пришлите темы для предмета **«{subj[1]}»** (каждая с новой строки).\n\n"
        "По умолчанию для каждой темы будет установлено **1 место**."
    )
    await state.set_state(AdminStates.waiting_for_topics)
    await call.answer()

@router.message(AdminStates.waiting_for_topics)
async def admin_add_topics_process(message: Message, state: FSMContext):
    if message.from_user.id != ADMIN_ID:
        return
    data = await state.get_data()
    list_id = data.get("target_list_id")
    
    lines = [line.strip() for line in message.text.split("\n") if line.strip()]
    if not lines:
        await message.reply("Пустой список.")
        return

    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        for topic in lines:
            cursor.execute("INSERT INTO topics (list_id, title, max_members) VALUES (?, ?, 1)", (list_id, topic))
        conn.commit()

    await state.clear()
    await message.reply(
        f"✅ Добавлено {len(lines)} тем!\n"
        "Студенты уже могут разбирать их в каталоге.",
        reply_markup=get_main_menu(message.from_user.id)
    )

@router.callback_query(F.data == "admin_limits_select")
async def admin_limits_select(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        return
    lists = get_all_lists()
    if not lists:
        await call.answer("Нет предметов.", show_alert=True)
        return
    buttons = [[InlineKeyboardButton(text=title, callback_data=f"adm_limlist_{l_id}")] for l_id, title in lists]
    buttons.append([InlineKeyboardButton(text="🔙 Назад", callback_data="admin_panel")])
    await call.message.edit_text("Выберите предмет для настройки количества мест:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await call.answer()

@router.callback_query(F.data.startswith("adm_limlist_"))
async def adm_limlist_topics(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        return
    list_id = int(call.data.split("_")[2])
    data = get_all_topics_report(list_id)
    if not data:
        await call.answer("В этом предмете пока нет тем!", show_alert=True)
        return

    buttons = []
    for t_id, title, max_m, members in data:
        display_title = title if len(title) <= 30 else title[:27] + "..."
        buttons.append([InlineKeyboardButton(text=f"⚙️ [{max_m} чел.] {display_title}", callback_data=f"editlim_{list_id}_{t_id}")])
    buttons.append([InlineKeyboardButton(text="🔙 Назад", callback_data="admin_limits_select")])
    await call.message.edit_text("Нажмите на тему для изменения лимита мест:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await call.answer()

@router.callback_query(F.data.startswith("editlim_"))
async def editlim_topic(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        return
    parts = call.data.split("_")
    list_id = int(parts[1])
    topic_id = int(parts[2])

    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT title, max_members FROM topics WHERE id = ?", (topic_id,))
        row = cursor.fetchone()

    title, max_m = row
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="1 чел.", callback_data=f"setlim_{list_id}_{topic_id}_1"),
            InlineKeyboardButton(text="2 чел.", callback_data=f"setlim_{list_id}_{topic_id}_2"),
            InlineKeyboardButton(text="3 чел.", callback_data=f"setlim_{list_id}_{topic_id}_3")
        ],
        [
            InlineKeyboardButton(text="4 чел.", callback_data=f"setlim_{list_id}_{topic_id}_4"),
            InlineKeyboardButton(text="✏️ Ввести другое число", callback_data=f"customlim_{list_id}_{topic_id}")
        ],
        [InlineKeyboardButton(text="🔙 Назад", callback_data=f"adm_limlist_{list_id}")]
    ])
    await call.message.edit_text(
        f"Тема: **«{title}»**\nТекущий лимит: **{max_m} чел.**\n\nВыберите новый лимит:",
        reply_markup=kb
    )
    await call.answer()

@router.callback_query(F.data.startswith("setlim_"))
async def setlim_fixed(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        return
    parts = call.data.split("_")
    list_id, topic_id, val = int(parts[1]), int(parts[2]), int(parts[3])
    update_topic_limit(topic_id, val)
    await call.answer(f"Лимит изменён на {val} чел.!", show_alert=True)
    await adm_limlist_topics(call)

@router.callback_query(F.data.startswith("customlim_"))
async def customlim_start(call: CallbackQuery, state: FSMContext):
    if call.from_user.id != ADMIN_ID:
        return
    parts = call.data.split("_")
    list_id, topic_id = int(parts[1]), int(parts[2])
    await state.update_data(target_list_id=list_id, target_topic_id=topic_id)
    await call.message.edit_text("Отправьте количество человек для этой темы:")
    await state.set_state(AdminStates.waiting_for_custom_limit)
    await call.answer()

@router.message(AdminStates.waiting_for_custom_limit)
async def customlim_finish(message: Message, state: FSMContext):
    if message.from_user.id != ADMIN_ID:
        return
    try:
        val = int(message.text.strip())
        if val < 1:
            raise ValueError
    except ValueError:
        await message.reply("Введите положительное число.")
        return

    data = await state.get_data()
    update_topic_limit(data["target_topic_id"], val)
    await state.clear()
    await message.reply(
        f"✅ Лимит успешно изменен на {val} чел.!",
        reply_markup=get_main_menu(message.from_user.id)
    )

@router.callback_query(F.data == "admin_delete_select")
async def admin_delete_select(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        return
    lists = get_all_lists()
    if not lists:
        await call.answer("Списков нет.", show_alert=True)
        return
    buttons = [[InlineKeyboardButton(text=f"🗑 {title}", callback_data=f"delconf_{l_id}")] for l_id, title in lists]
    buttons.append([InlineKeyboardButton(text="🔙 Назад", callback_data="admin_panel")])
    await call.message.edit_text("Какой предмет удалить?", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await call.answer()

@router.callback_query(F.data.startswith("delconf_"))
async def delconf_handler(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        return
    list_id = int(call.data.split("_")[1])
    subj = get_list_by_id(list_id)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⚠️ Да, удалить безвозвратно", callback_data=f"delexec_{list_id}")],
        [InlineKeyboardButton(text="❌ Отмена", callback_data="admin_delete_select")]
    ])
    await call.message.edit_text(f"Удалить предмет **«{subj[1]}»** и все темы в нём?", reply_markup=kb)
    await call.answer()

@router.callback_query(F.data.startswith("delexec_"))
async def delexec_handler(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        return
    list_id = int(call.data.split("_")[1])
    delete_list(list_id)
    await call.answer("Предмет удален!", show_alert=True)
    await admin_panel(call)

@router.callback_query(F.data == "admin_view_select")
async def admin_view_select(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        return
    lists = get_all_lists()
    if not lists:
        await call.answer("Списков нет.", show_alert=True)
        return
    buttons = [[InlineKeyboardButton(text=title, callback_data=f"adm_view_{l_id}")] for l_id, title in lists]
    buttons.append([InlineKeyboardButton(text="🔙 Назад", callback_data="admin_panel")])
    await call.message.edit_text("Какой предмет посмотреть текстом?", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await call.answer()

@router.callback_query(F.data.startswith("adm_view_"))
async def adm_view_exec(call: CallbackQuery):
    if call.from_user.id != ADMIN_ID:
        return
    list_id = int(call.data.split("_")[2])
    subj = get_list_by_id(list_id)
    data = get_all_topics_report(list_id)
    
    text = f"📋 **Реестр: «{subj[1]}»**\n\n"
    for idx, (t_id, title, max_m, members) in enumerate(data, start=1):
        cur_cnt = len(members)
        if cur_cnt == 0:
            text += f"🟢 {idx}. {title} — (0/{max_m})\n"
        elif cur_cnt < max_m:
            text += f"🟡 {idx}. {title} — 👤 {', '.join(members)} [{cur_cnt}/{max_m}]\n"
        else:
            text += f"❌ {idx}. {title} — 👤 {', '.join(members)} [Заполнена]\n"

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🔙 Назад", callback_data="admin_view_select")]
    ])
    await call.message.edit_text(text, reply_markup=kb)
    await call.answer()

async def main():
    init_db()
    await bot.delete_webhook(drop_pending_updates=True)
    print("Бот успешно запущен: мульти-предметный режим активен!")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())