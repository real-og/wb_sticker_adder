from loader import dp
from aiogram import types
from aiogram.dispatcher import FSMContext
import texts

@dp.message_handler(commands=['start'], state="*")
async def send_welcome(message: types.Message, state: FSMContext):
    await message.answer(f'{message.from_user.language_code} - {texts.your_lang}')