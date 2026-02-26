import asyncio
import logging
import pandas as pd
from aiogram import Bot, Dispatcher, types
from aiogram.types import Message
from aiogram.filters import Command
from aiogram import F
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties

TOKEN = "8424122720:AAG8gR5D119GJKXFowDSBHD4vOOoDMAX2KI"

logging.basicConfig(level=logging.INFO)

bot = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher()

user_words = {}

@dp.message(Command("start"))
async def start_handler(message: Message):
    await message.answer(
        "Отправь слова через запятую.\n"
        "Потом отправь Excel файл (.xlsx)"
    )

@dp.message(F.document)
async def handle_document(message: Message):
    if not message.document.file_name.endswith(".xlsx"):
        await message.answer("Нужен файл формата .xlsx")
        return

    if message.from_user.id not in user_words:
        await message.answer("Сначала отправь список слов.")
        return

    file = await bot.get_file(message.document.file_id)
    downloaded_file = await bot.download_file(file.file_path)

    df = pd.read_excel(downloaded_file)

    words = user_words[message.from_user.id]
    results = []

    # Индексы столбцов по буквам Excel
    col_F = 5     # F
    col_AI = 34   # AI
    col_AK = 36   # AK

    # Общий фильтр AI > 0
    df_positive = df[df.iloc[:, col_AI] > 0]

    for word in words:
        filtered = df[
            (df.iloc[:, col_F] == word) &
            (df.iloc[:, col_AI] > 0)
        ]

        if filtered.empty:
            results.append([word, None, None, None, None])
        else:
            results.append([
                word,
                filtered.iloc[:, col_AK].max(),
                filtered.iloc[:, col_AK].min(),
                filtered.iloc[:, col_AK].mean(),
                filtered.iloc[:, col_AI].sum()
            ])

    # Добавляем общие показатели
    total_sum_AI = df_positive.iloc[:, col_AI].sum()
    total_avg_AK = df_positive.iloc[:, col_AK].mean()

    results.append(["", "", "", "", ""])
    results.append(["ИТОГО", "", "", "", ""])
    results.append(["Общая сумма AI", "", "", "", total_sum_AI])
    results.append(["Общее среднее AK", "", "", total_avg_AK, ""])

    result_df = pd.DataFrame(
        results,
        columns=["Слово", "Макс.", "Мин.", "Среднее.", "Сумма"]
    )

    output_file = "result.xlsx"
    result_df.to_excel(output_file, index=False)

    await message.answer_document(types.FSInputFile(output_file))


@dp.message()
async def get_words(message: Message):
    words = [w.strip() for w in message.text.split(",")]
    user_words[message.from_user.id] = words
    await message.answer("Слова сохранены. Теперь отправь Excel файл.")


async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
