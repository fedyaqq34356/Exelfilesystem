# bot.py
import asyncio
import logging
import hashlib
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.enums import ParseMode

import config
from database import Database
from excel_handler import ExcelHandler
from file_monitor import FileMonitor

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

bot = Bot(token=config.BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())
db = Database()
excel = ExcelHandler()
monitor = FileMonitor()

active_applications = {}

class SettingsState(StatesGroup):
    waiting_path = State()


@dp.message(Command("start"))
async def cmd_start(message: Message):
    await message.answer(
        "Вітаю!\n\n"
        "Я бот для автоматичного погодження заявок на оплату.\n\n"
        "<b>Доступні команди:</b>\n"
        "/status — поточний статус системи\n"
        "/settings — налаштування папок\n"
        "/stats — статистика обробки\n"
        "/help — детальна допомога",
        parse_mode=ParseMode.HTML
    )


@dp.message(Command("help"))
async def cmd_help(message: Message):
    help_text = (
        "ДОВІДКА ПО СИСТЕМІ\n\n"
        "<b>Як це працює:</b>\n"
        "1. Файли .xlsm з'являються в папках Фіндиректора або Директора\n"
        "2. Бот автоматично читає дані заявки\n"
        "3. Відправляє повідомлення відповідальній особі\n"
        "4. Після натискання кнопки файл переміщується далі\n\n"
        "<b>Команди:</b>\n"
        "/status — статус папок та моніторингу\n"
        "/settings — змінити шляхи до папок\n"
        "/stats — переглянути статистику\n\n"
        "<b>Налаштування в .env:</b>\n"
        "BOT_TOKEN — токен бота\n"
        "CHAT_ID_FINDIRECTOR — ID чату фіндиректора\n"
        "CHAT_ID_DIRECTOR — ID чату директора\n"
        "CHECK_INTERVAL — інтервал перевірки (секунди)"
    )
    await message.answer(help_text, parse_mode=ParseMode.HTML)


@dp.message(Command("status"))
async def cmd_status(message: Message):
    stats = monitor.get_monitoring_stats()
    
    text = "СТАТУС СИСТЕМИ\n\n"
    text += f"Остання перевірка: {stats['last_check']}\n"
    text += f"Інтервал: {config.CHECK_INTERVAL} сек\n"
    text += f"Оброблено всього: {len(db.get_all_processed_files())} файлів\n\n"
    
    text += "<b>СТАТУС ПАПОК:</b>\n"
    for name, info in stats['folders'].items():
        emoji = "Активно" if info["exists"] else "Не знайдено"
        status = "активна" if info["exists"] else "не знайдена"
        text += f"{emoji} <b>{name}</b>: {status}, файлів: {info['files']}\n"
    
    await message.answer(text, parse_mode=ParseMode.HTML)


@dp.message(Command("stats"))
async def cmd_stats(message: Message):
    recent = db.get_recent_actions(20)
    
    text = "ОСТАННЯ АКТИВНІСТЬ\n\n"
    if not recent:
        text += "Поки що заявок не оброблено"
    else:
        for action in recent:
            emoji = "Погоджено" if action['action'] == "APPROVED" else "Відхилено" if action['action'] == "REJECTED" else "Виявлено"
            text += f"{emoji} <code>{action['file_name']}</code>\n"
            text += f"   {action['action']} by {action['user']}\n"
            text += f"   {action['timestamp']}\n\n"
    
    await message.answer(text, parse_mode=ParseMode.HTML)


@dp.message(Command("settings"))
async def cmd_settings(message: Message):
    kb = [
        [InlineKeyboardButton(text="Папка Фіндиректора", callback_data="set_findirector_folder")],
        [InlineKeyboardButton(text="Папка Директора", callback_data="set_director_folder")],
        [InlineKeyboardButton(text="Папка Бухгалтера", callback_data="set_accountant_folder")],
        [InlineKeyboardButton(text="Папка Касира", callback_data="set_cashier_folder")],
        [InlineKeyboardButton(text="Папка Відхилених", callback_data="set_rejected_folder")],
        [InlineKeyboardButton(text="Закрити", callback_data="close_settings")],
    ]
    await message.answer(
        "НАЛАШТУВАННЯ ПАПОК\n\nОберіть папку для зміни:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb),
        parse_mode=ParseMode.HTML
    )


@dp.callback_query(F.data == "close_settings")
async def close_settings(cb: CallbackQuery):
    await cb.message.delete()
    await cb.answer()


@dp.callback_query(F.data.startswith("set_"))
async def set_folder(cb: CallbackQuery, state: FSMContext):
    key = cb.data.replace("set_", "")
    current = config.get_path(key) or "Не налаштовано"
    
    folder_names = {
        "findirector_folder": "Фіндиректора",
        "director_folder": "Директора",
        "accountant_folder": "Бухгалтера",
        "cashier_folder": "Касира",
        "rejected_folder": "Відхилених"
    }
    
    await state.update_data(key=key)
    await state.set_state(SettingsState.waiting_path)
    await cb.message.edit_text(
        f"Папка {folder_names.get(key, key)}\n\n"
        f"Поточний шлях:\n<code>{current}</code>\n\n"
        f"Надішліть новий шлях:",
        parse_mode=ParseMode.HTML
    )
    await cb.answer()


@dp.message(SettingsState.waiting_path)
async def save_new_path(message: Message, state: FSMContext):
    data = await state.get_data()
    key = data["key"]
    new_path = message.text.strip()
    
    if config.update_path(key, new_path):
        await message.answer(
            f"Шлях оновлено!\n\n<code>{new_path}</code>",
            parse_mode=ParseMode.HTML
        )
        logger.info(f"Шлях {key} оновлено на: {new_path}")
    else:
        await message.answer("Помилка оновлення шляху")
    
    await state.clear()


async def send_application(data):
    if data["intended_approver"] == "ФІНДИРЕКТОР":
        chat_id = config.CHAT_ID_FINDIRECTOR
    elif data["intended_approver"] == "ДИРЕКТОР":
        chat_id = config.CHAT_ID_DIRECTOR
    else:
        logger.error(f"Невідомий погоджувач: {data['intended_approver']}")
        return

    if not chat_id:
        logger.error(f"Chat ID не налаштовано для {data['intended_approver']}")
        return

    file_id = hashlib.md5(data['file_path'].encode('utf-8')).hexdigest()

    text = (
        "НОВА ЗАЯВКА НА ПОГОДЖЕННЯ\n\n"
        f"Дата: <b>{data['дата']}</b>\n"
        f"Заявник: {data['заявник']}\n"
        f"Відділ: {data['відділ']}\n"
        f"Сума: <b>{data['сума']}</b>\n"
        f"Постачальник: {data['постачальник']}\n"
        f"Вид розрахунку: {data['вид_розрахунку']}\n\n"
        f"<b>Призначення:</b>\n{data['призначення']}\n\n"
        f"Файл: <code>{data['file_name']}</code>\n"
        f"Погоджує: <b>{data['intended_approver']}</b>"
    )

    kb = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="ПОГОДИТИ", callback_data=f"approve_{file_id}"),
        InlineKeyboardButton(text="ВІДХИЛИТИ", callback_data=f"reject_{file_id}")
    ]])

    try:
        active_applications[file_id] = data
        await bot.send_message(chat_id, text, reply_markup=kb, parse_mode=ParseMode.HTML)
        logger.info(f"Заявка відправлена: {data['file_name']} → {data['intended_approver']}")
    except Exception as e:
        logger.error(f"Помилка відправки: {e}")


@dp.callback_query(F.data.startswith("approve_"))
async def approve(cb: CallbackQuery):
    await cb.answer()
    
    file_id = cb.data[len("approve_"):]
    data = active_applications.get(file_id)
    
    if not data:
        await cb.message.answer("Заявка вже оброблена або не знайдена")
        return

    await cb.message.edit_text("Обробка заявки...")

    success = excel.move_file(data['file_path'], approved=True)
    
    if success:
        user_name = cb.from_user.first_name or cb.from_user.username or "Невідомо"
        db.log_action(data['file_name'], "APPROVED", user_name, f"Сума: {data['сума']}")
        
        await cb.message.edit_text(
            f"ЗАЯВКУ ПОГОДЖЕНО\n\n"
            f"{data['file_name']}\n"
            f"{data['сума']}\n"
            f"Погодив: {user_name}\n"
            f"Файл переміщено далі по маршруту",
            parse_mode=ParseMode.HTML
        )
        logger.info(f"APPROVED: {data['file_name']} by {user_name}")
    else:
        await cb.message.edit_text("Помилка переміщення файлу")

    active_applications.pop(file_id, None)


@dp.callback_query(F.data.startswith("reject_"))
async def reject(cb: CallbackQuery):
    await cb.answer()
    
    file_id = cb.data[len("reject_"):]
    data = active_applications.get(file_id)
    
    if not data:
        await cb.message.answer("Заявка вже оброблена або не знайдена")
        return

    await cb.message.edit_text("Відхилення заявки...")

    success = excel.move_file(data['file_path'], approved=False)
    
    if success:
        user_name = cb.from_user.first_name or cb.from_user.username or "Невідомо"
        db.log_action(data['file_name'], "REJECTED", user_name, f"Сума: {data['сума']}")
        
        await cb.message.edit_text(
            f"ЗАЯВКУ ВІДХИЛЕНО\n\n"
            f"{data['file_name']}\n"
            f"{data['сума']}\n"
            f"Відхилив: {user_name}\n"
            f"Файл переміщено в папку «Відхилені»",
            parse_mode=ParseMode.HTML
        )
        logger.info(f"REJECTED: {data['file_name']} by {user_name}")
    else:
        await cb.message.edit_text("Помилка переміщення файлу")

    active_applications.pop(file_id, None)


async def monitoring_task():
    logger.info("Моніторинг розпочато")
    
    while True:
        try:
            new_applications = monitor.check_folders()
            
            for app in new_applications:
                await send_application(app)
                await asyncio.sleep(0.5)
            
            await asyncio.sleep(config.CHECK_INTERVAL)
            
        except Exception as e:
            logger.error(f"Помилка моніторингу: {e}", exc_info=True)
            await asyncio.sleep(10)


async def main():
    logger.info("="*70)
    logger.info("ЗАПУСК БОТА ДЛЯ ПОГОДЖЕННЯ ЗАЯВОК")
    logger.info("="*70)
    
    if not config.BOT_TOKEN or config.BOT_TOKEN == "YOUR_BOT_TOKEN_HERE":
        logger.error("BOT_TOKEN не налаштовано в .env файлі!")
        return
    
    config.ensure_folders_exist()
    
    logger.info("Перевірка папок:")
    for key, path in config.PATHS.items():
        if path:
            from pathlib import Path
            exists = Path(path).exists()
            status = "Активно" if exists else "Не знайдено"
            logger.info(f"   {status} {key}: {path}")
    
    asyncio.create_task(monitoring_task())
    
    logger.info("="*70)
    logger.info("БОТ ПРАЦЮЄ")
    logger.info("="*70)
    
    await dp.start_polling(bot, allowed_updates=dp.resolve_used_update_types())


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Зупинка бота...")
    except Exception as e:
        logger.error(f"Критична помилка: {e}", exc_info=True)