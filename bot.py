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
from openpyxl.styles import PatternFill, Font, Alignment

TOKEN = os.getenv("BOT_TOKEN")

logging.basicConfig(level=logging.INFO)

bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher()

user_words = {}

# ---------------- ГРУППЫ ---------------- #

groups = {
    "Пальто": [
        "drap_2-beige","drap_2-black","drap_2-grey",
        "drap-beige","drap-brown","drap-grey","drap-black",
        "coat-jacket-beige","coat-jacket-grey","coat-jacket-brown",
        "drap_3(belt)-grey","drap_3(belt)-brown",
        "drap_3(belt)-grafit","drap_3(belt)-beige"
    ],
    "Куртки": [
        "suede-ohra","suede-milk","suede-brown",
        "bomber-ohra","bomber-brown","bomber-milk"
    ],
    "двойка топ с завязками": [
        "bows-blue","bows-cappuccino","bows-haki"
    ],
    "Песок": [
        "costum-black","costum-blue","costum-brown",
        "costum-green","costum-grey","costum-olive"
    ],
    "Шанель": [
        "flax-beige","flax-blue","flax-brown",
        "flax-green","flax-grey"
    ],
    "Шорты": [
        "short1beige","short1blue","short1fuksia",
        "short1green","short-haki","short-brown",
        "short-black","short-mentol23"
    ]
}

group_colors = {
    "Пальто": "FFF2CC",
    "Куртки": "F8CBAD",
    "двойка топ с завязками": "BDD7EE",
    "Песок": "FCE4D6",
    "Шанель": "E2EFDA",
    "Шорты": "D9E1F2"
}

# ---------------- СТАРТ ---------------- #

@dp.message(Command("start"))
async def start_handler(message: Message):
    await message.answer(
        "Отправь слова через запятую.\n"
        "Потом отправь Excel файл (.xlsx)"
    )

# ---------------- ПРИЁМ СЛОВ ---------------- #

@dp.message(F.text)
async def save_words(message: Message):
    words = [w.strip() for w in message.text.split(",")]
    user_words[message.from_user.id] = words
    await message.answer("Слова сохранены. Теперь отправь Excel файл.")

# ---------------- ПРИЁМ ФАЙЛА ---------------- #

@dp.message(F.document)
async def handle_document(message: Message):

    if message.from_user.id not in user_words:
        await message.answer("Сначала отправь список слов.")
        return

    if not message.document.file_name.endswith(".xlsx"):
        await message.answer("Нужен файл формата .xlsx")
        return

    await message.answer("Файл получен. Обработка началась...")

    file = await bot.get_file(message.document.file_id)
    downloaded_file = await bot.download_file(file.file_path)

    df = pd.read_excel(downloaded_file)

    col_F = 5
    col_AI = 34
    col_AK = 36

    results_by_group = {}
    avg_values = []
    total_sum = 0

    for group_name, group_words in groups.items():
        results_by_group[group_name] = []

        for word in group_words:

            filtered = df[
                (df.iloc[:, col_F] == word) &
                (df.iloc[:, col_AI] > 0)
            ]

            if filtered.empty:
                max_val = min_val = avg_val = sum_val = 0
            else:
                max_val = filtered.iloc[:, col_AK].max()
                min_val = filtered.iloc[:, col_AK].min()
                avg_val = filtered.iloc[:, col_AK].mean()
                sum_val = filtered.iloc[:, col_AI].sum()

                avg_values.append(avg_val)
                total_sum += sum_val

            results_by_group[group_name].append(
                [word, max_val, min_val, avg_val, sum_val]
            )

    average_of_averages = (
        sum(avg_values) / len(avg_values)
        if avg_values else 0
    )

    # -------- СОЗДАНИЕ EXCEL -------- #

    wb = Workbook()
    ws = wb.active

    yesterday = datetime.now() - timedelta(days=1)
    date_text = yesterday.strftime("%d.%m.%Y")

    ws.merge_cells("A1:E1")
    ws["A1"] = f"За {date_text}"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="right")

    ws.append(["Слово", "Макс.", "Мин.", "Среднее", "Доставки шт."])

    current_row = 3

    for group_name, rows in results_by_group.items():

        ws.cell(row=current_row, column=1, value=group_name)
        ws.cell(row=current_row, column=1).font = Font(bold=True)
        current_row += 1

        fill = PatternFill(
            start_color=group_colors[group_name],
            end_color=group_colors[group_name],
            fill_type="solid"
        )

        for row_data in rows:
            for col_index, value in enumerate(row_data, start=1):
                cell = ws.cell(row=current_row, column=col_index, value=value)
                cell.fill = fill
                cell.alignment = Alignment(horizontal="center")
            current_row += 1

    current_row += 1
    ws.cell(row=current_row, column=1, value="Всего доставок шт.")
    ws.cell(row=current_row, column=5, value=total_sum)

    current_row += 1
    ws.cell(row=current_row, column=1, value="Среднее от среднего")
    ws.cell(row=current_row, column=4, value=average_of_averages)

    for column in ws.columns:
        max_length = 0
        col_letter = column[0].column_letter
        for cell in column:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

    file_name = "result.xlsx"
    wb.save(file_name)

    await message.answer_document(types.FSInputFile(file_name))

# ---------------- ЗАПУСК ---------------- #

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
