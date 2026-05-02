# Полный прогон pipeline для нового Word-файла

Эта инструкция описывает ручной прогон WordKeywords для нового сборника статей из PowerShell. Пример использует `input\test7.docx` и run-tag `test7`.

## Ключевое ограничение

Основной текст заказчика считается эталоном. Extraction/debug scripts читают input `.docx`, а финальный builder делает byte-copy в `output\` и добавляет новые разделы только после последней статьи.

## Полный порядок команд

Запускать из корня репозитория:

```powershell
python dump_doc_paragraph_snapshot.py --docx input\test7.docx --run-tag test7
python build_author_title_paragraph_ru_from_snapshot.py --run-tag test7
python scripts/enrich_ru_title_paragraph_structure.py --docx input\test7.docx --run-tag test7
python build_draft_author_index_ru_from_snapshot.py --run-tag test7
python build_author_index_ru_text_from_snapshot.py --run-tag test7

python debug_toc_ru_from_word.py --docx input\test7.docx --run-tag test7
python build_toc_ru_draft_text.py --run-tag test7

python scripts/debug_author_windows_en.py --run-tag test7
python -m scripts.debug_en_title_author_pairs --docx input\test7.docx --run-tag test7
python -m scripts.build_draft_author_index_en --docx input\test7.docx --run-tag test7
python -m scripts.build_draft_toc_en --docx input\test7.docx --run-tag test7

python -m scripts.fast_keyword_index_find --docx input\test7.docx --run-tag test7
python -m scripts.build_separate_keyword_indexes --run-tag test7

python build_final_document.py --docx input\test7.docx --run-tag test7
```

## Основные artifacts

- `output\test7_doc_paragraph_snapshot.csv`
- `output\test7_author_title_paragraph_ru_debug_from_snapshot.csv`
- `output\test7_ru_title_paragraph_structure_debug.csv`
- `output\test7_author_index_ru_from_snapshot.txt`
- `output\test7_author_index_ru_from_snapshot.csv`
- `output\test7_draft_author_index_en.txt`
- `output\test7_draft_author_index_en.csv`
- `output\test7_keyword_index_ru.txt`
- `output\test7_keyword_index_ru.csv`
- `output\test7_keyword_index_en.txt`
- `output\test7_keyword_index_en.csv`
- `output\test7_draft_toc_ru.csv`
- `output\test7_draft_toc_ru.txt`
- `output\test7_draft_toc_en.csv`
- `output\test7_draft_toc_en.txt`
- `output\test7_final_ordered.docx`

## Итоговая структура документа

1. Основной текст статей
2. `Авторский указатель`
3. `Author Index`
4. `Предметный указатель`
5. `Keyword Index`
6. `Оглавление`
7. `Table of Contents`

## Контрольные признаки успешного прогона

- Итоговый файл создан: `output\test7_final_ordered.docx`.
- Основной текст не изменён.
- Добавленные разделы идут только после статей.
- Нумерация страниц appendices сквозная.
- RU/EN TOC в одну колонку.
- Keyword formulas сохраняют hyphen и subscript.
- EN TOC контрольные pages: `117`, `180`, `196`, `223`, `232`, `312`, `419`.
- В RU/EN TOC страницы указателей соответствуют реальным страницам.

## Типовые ошибки

- `FileNotFoundError` по промежуточному CSV обычно означает, что пропущен предыдущий шаг или указан неверный `--run-tag`.
- Если Word COM падает с `Вызов был отклонен`, закрыть Word/убить `WINWORD` и повторить.
- Если `pywin32` `gen_py` cache сломан, очистить путь из `win32com.client.gencache.GetGeneratePath()`.
- Если `NOT FOUND > 0`, смотреть соответствующие debug `.txt` и `.csv`.
- Если input `.docx` имеет размер `0`, файл скопирован неправильно.
- Если скрипт ищет не те tagged artifacts, проверить единый `--run-tag` во всех командах.
- Если heading не найден при расчёте страниц указателей, смотреть diagnostic output `build_final_document.py`.

## Что не коммитить

- `output\*.docx`
- Временные Word-файлы, включая `~$*.docx`
- Большие промежуточные артефакты из `output\`, если они покрыты `.gitignore`
- Локальные ручные копии итоговых документов
