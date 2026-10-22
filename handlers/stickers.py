"""Добавление артикула продавца на стикеры WB (aiogram 2.25.1)."""
import asyncio
import logging
import re
import tempfile
from pathlib import Path

import fitz
from aiogram import types
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from openpyxl import load_workbook

from loader import bot, dp

LOG = logging.getLogger(__name__)
if not any(isinstance(h, logging.FileHandler) for h in LOG.handlers):
    handler = logging.FileHandler('sticker_errors.log', encoding='utf-8')
    handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
    LOG.addHandler(handler)
LOG.setLevel(logging.ERROR)

MAX_BYTES = 20 * 1024 * 1024


class StickerFiles(StatesGroup):
    waiting = State()


def parse_excel(path):
    try:
        book = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise ValueError('Не удалось открыть XLSX: файл повреждён или имеет неверный формат.') from exc
    try:
        sheet = book.active
        sheet.reset_dimensions()  # Экспорт WB может ошибочно указывать A1:A1.
        mapping = {}
        for row_num, (article, number) in enumerate(
            sheet.iter_rows(min_row=6, min_col=3, max_col=4, values_only=True), 6
        ):
            if article is None and number is None:
                continue
            if article is None or number is None:
                raise ValueError(f'Excel, строка {row_num}: не заполнен артикул или номер стикера.')
            if isinstance(number, bool):
                raise ValueError(f'Excel, строка {row_num}: неверный номер стикера.')
            if isinstance(number, (int, float)):
                if int(number) != number:
                    raise ValueError(f'Excel, строка {row_num}: номер не целый.')
                number = str(int(number))
            key = re.sub(r'\s+', '', str(number))
            if not re.fullmatch(r'\d{11}', key):
                raise ValueError(f'Excel, строка {row_num}: неверный номер {number!r} (ожидается 11 цифр).')
            value = str(article).strip()
            if not value or not value.isascii() or not value.isprintable():
                raise ValueError(f'Excel, строка {row_num}: артикул пустой или содержит неподдерживаемые символы.')
            if key in mapping:
                raise ValueError(f'Excel, строка {row_num}: повторяется номер {key}.')
            mapping[key] = value
        if not mapping:
            raise ValueError('В Excel нет строк с товарами (ожидаются столбцы C и D с 6-й строки).')
        return mapping
    finally:
        book.close()


def make_stickers(pdf_path, xlsx_path, output_path):
    mapping = parse_excel(xlsx_path)
    try:
        pdf = fitz.open(pdf_path)
    except Exception as exc:
        raise ValueError('Не удалось открыть PDF: файл повреждён или имеет неверный формат.') from exc
    try:
        if pdf.needs_pass:
            raise ValueError('PDF защищён паролем.')
        if not pdf.page_count:
            raise ValueError('PDF не содержит страниц.')
        found = []
        for index, page in enumerate(pdf, 1):
            if page.rotation != 0 or not (110 <= page.rect.width <= 116 and 82 <= page.rect.height <= 88):
                raise ValueError(f'Страница {index}: размер или ориентация стикера отличаются от поддерживаемых 40×30 мм.')
            lines = [s.strip() for s in page.get_text('text', sort=False).splitlines()]
            ids = [a + b for a, b in zip(lines, lines[1:])
                   if re.fullmatch(r'\d{7}', a) and re.fullmatch(r'\d{4}', b)]
            if len(ids) != 1:
                raise ValueError(f'Страница {index}: не удалось однозначно прочитать номер стикера.')
            key = ids[0]
            if key not in mapping:
                raise ValueError(f'Страница {index}: стикер {key} отсутствует в Excel.')
            if key in found:
                raise ValueError(f'Страница {index}: номер {key} повторяется в PDF.')
            found.append(key)
        extra = mapping.keys() - set(found)
        if extra:
            raise ValueError(f'В Excel {len(mapping)} товаров, в PDF {len(found)} стикеров. '
                             f'Нет стикеров для {len(extra)} номеров, например {next(iter(sorted(extra)))}.')

        # Сначала проверяем, что все надписи помещаются в узкой белой зоне слева от WB.
        sizes = []
        for key in found:
            article = mapping[key]
            width = fitz.get_text_length(article, fontname='hebo', fontsize=1)
            size = min(3.8, 30 / max(width, 0.01))
            if size < 2.5:
                raise ValueError(f'Артикул {article!r} слишком длинный для свободного поля стикера.')
            sizes.append(size)
        for page, key, size in zip(pdf, found, sizes):
            page.insert_text((7.0, 58), mapping[key], fontname='hebo',
                             fontsize=size, rotate=90, color=(0, 0, 0), overlay=True)
        pdf.save(output_path, garbage=3, deflate=True)
        return len(found)
    finally:
        pdf.close()


@dp.message_handler(commands=['cancel'], state='*')
async def cancel_stickers(message: types.Message, state: FSMContext):
    await state.finish()
    await message.answer('Загрузка отменена. Отправьте PDF и XLSX заново.')


@dp.message_handler(content_types=types.ContentType.DOCUMENT, state='*')
async def sticker_document(message: types.Message, state: FSMContext):
    ext = Path(message.document.file_name or '').suffix.lower()
    if ext not in ('.pdf', '.xlsx'):
        return await message.answer('Нужны два документа: PDF со стикерами и XLSX с артикулами.')
    if message.document.file_size and message.document.file_size > MAX_BYTES:
        return await message.answer('Файл больше 20 МБ. Отправьте файл меньшего размера.')
    data = await state.get_data()
    if ext in data:
        return await message.answer(f'{ext.upper()} уже получен. Для замены используйте /cancel.')
    await state.update_data(**{ext: message.document.file_id, f'{ext}_name': message.document.file_name or 'stickers.pdf'})
    await StickerFiles.waiting.set()
    data = await state.get_data()
    if '.pdf' not in data or '.xlsx' not in data:
        return await message.answer(f'Файл получен. Теперь отправьте {"XLSX" if ext == ".pdf" else "PDF"}.')

    try:
        await message.answer('Проверяю номера и формирую PDF...')
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path, xlsx_path, result = (str(Path(tmp) / name) for name in
                                           ('input.pdf', 'input.xlsx', 'ready.pdf'))
            await bot.download_file_by_id(data['.pdf'], pdf_path)
            await bot.download_file_by_id(data['.xlsx'], xlsx_path)
            count = await asyncio.to_thread(make_stickers, pdf_path, xlsx_path, result)
            with open(result, 'rb') as file:
                await message.answer_document(
                    types.InputFile(file, filename=f"{Path(data['.pdf_name']).stem}_added.pdf"),
                    caption=f'Готово: {count} из {count} стикеров.'
                )
    except ValueError as exc:
        LOG.exception('user_id=%s: ошибка данных', message.from_user.id)
        await message.answer(f'Не удалось создать PDF. Причина: {exc}\nОтправьте исправленные файлы заново.')
    except Exception:
        LOG.exception('user_id=%s: системная ошибка', message.from_user.id)
        await message.answer('Не удалось создать или отправить PDF из-за технической ошибки. '
                             'Подробности записаны в sticker_errors.log. Попробуйте ещё раз.')
    finally:
        await state.finish()
