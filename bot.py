import asyncio
import logging
import os
import pandas as pd
from datetime import datetime, timedelta
from aiogram import Bot, Dispatcher, types
from aiogram.types import Message
from aiogram.filters import Command
from aiogram import F
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

TOKEN = os.getenv("BOT_TOKEN")

logging.basicConfig(level=logging.INFO)

bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher()

user_words = {}

@dp.message(Command("start"))
async def start_handler(message: Message):
    await message.answer("Отправь слова через запятую, затем Excel файл.")

@dp.message(F.document)
async def handle_document(message: Message):
    if not message.document.file_name.endswith(".xlsx"):
        await message.answer("Нужен файл .xlsx")
        return

    if message.from_user.id not in user_words:
        await message.answer("Сначала отправь список слов.")
        return

    file = await bot.get_file(message.document.file_id)
    downloaded_file = await bot.download_file(file.file_path)

    df = pd.read_excel(downloaded_file)

    col_F = 5
    col_AI = 34
    col_AK = 36

    words = user_words[message.from_user.id]

    results = []
    avg_values = []

    for word in words:
        filtered = df[
            (df.iloc[:, col_F] == word) &
            (df.iloc[:, col_AI] > 0)
        ]

        if filtered.empty:
            max_val = min_val = avg_val = 0
            sum_val = 0
        else:
            max_val = filtered.iloc[:, col_AK].max()
            min_val = filtered.iloc[:, col_AK].min()
            avg_val = filtered.iloc[:, col_AK].mean()
            sum_val = filtered.iloc[:, col_AI].sum()

        results.append([word, max_val, min_val, avg_val, sum_val])
        avg_values.append(avg_val)

    # Общие показатели
    total_sum = sum([r[4] for r in results])
    average_of_averages = sum(avg_values) / len(avg_values) if avg_values else 0

    # Создание красивого Excel
    wb = Workbook()
    ws = wb.active

    # 📅 Дата вчера
    yesterday = datetime.now() - timedelta(days=1)
    date_text = yesterday.strftime("За %d %B")

    ws["A1"] = date_text
    ws["A1"].font = Font(size=14, bold=True)

    headers = ["", "Max", "Min", "Average", "Доставки шт."]
    ws.append(headers)

    row_num = 3
    for row in results:
        ws.append(row)

    ws.append(["", "", "", "", ""])
    ws.append(["Всего доставок шт.", "", "", "", total_sum])
    ws.append(["Среднее от среднего", "", "", average_of_averages, ""])

    # Выравнивание
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center")

    file_name = "report.xlsx"
    wb.save(file_name)

    await message.answer_document(types.FSInputFile(file_name))

@dp.message()
async def get_words(message: Message):
    words = [w.strip() for w in message.text.split(",")]
    user_words[message.from_user.id] = words
    await message.answer("Слова сохранены. Теперь отправь Excel файл.")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
