# WordKeywords / WordIndexer

## Назначение проекта

WordKeywords / WordIndexer — pipeline для Word-документа со статьями. Он извлекает RU/EN keywords и RU/EN authors, строит RU/EN author indexes, RU/EN keyword indexes, RU/EN TOC (оглавления) и собирает итоговый ordered `.docx`.

## Входной и итоговый документы

Pipeline параметризован через `--docx` и `--run-tag`.

Пример:

- Input: `input/test7.docx`
- Run tag: `test7`
- Output: `output/test7_final_ordered.docx`

Для другого файла используйте свой input `.docx` и тот же `run-tag` во всех командах pipeline.

## Важное ограничение

Исходный основной текст заказчика нельзя менять. Скрипты могут читать input `.docx`, но финальный builder должен сначала сделать byte-copy в `output/`, открыть только output-копию и добавить новые разделы только после последней статьи.

## Итоговая структура output docx

1. Основной текст статей
2. Авторский указатель
3. Author Index
4. Предметный указатель
5. Keyword Index
6. Оглавление
7. Table of Contents

## Полный pipeline для нового файла

Пример для `input/test7.docx` и `--run-tag test7`:

```powershell
python dump_doc_paragraph_snapshot.py --docx input/test7.docx --run-tag test7
python build_author_title_paragraph_ru_from_snapshot.py --run-tag test7
python scripts/enrich_ru_title_paragraph_structure.py --docx input/test7.docx --run-tag test7
python build_draft_author_index_ru_from_snapshot.py --run-tag test7
python build_author_index_ru_text_from_snapshot.py --run-tag test7

python debug_toc_ru_from_word.py --docx input/test7.docx --run-tag test7
python build_toc_ru_draft_text.py --run-tag test7

python scripts/debug_author_windows_en.py --run-tag test7
python -m scripts.debug_en_title_author_pairs --docx input/test7.docx --run-tag test7
python -m scripts.build_draft_author_index_en --docx input/test7.docx --run-tag test7
python -m scripts.build_draft_toc_en --docx input/test7.docx --run-tag test7

python -m scripts.fast_keyword_index_find --docx input/test7.docx --run-tag test7
python -m scripts.build_separate_keyword_indexes --run-tag test7

python build_final_document.py --docx input/test7.docx --run-tag test7
```

## Важные output artifacts

- `output/<run-tag>_doc_paragraph_snapshot.csv`
- `output/<run-tag>_author_index_ru_from_snapshot.txt`
- `output/<run-tag>_author_index_ru_from_snapshot.csv`
- `output/<run-tag>_draft_author_index_en.txt`
- `output/<run-tag>_draft_author_index_en.csv`
- `output/<run-tag>_keyword_index_ru.txt`
- `output/<run-tag>_keyword_index_ru.csv`
- `output/<run-tag>_keyword_index_en.txt`
- `output/<run-tag>_keyword_index_en.csv`
- `output/<run-tag>_draft_toc_ru.csv`
- `output/<run-tag>_draft_toc_ru.txt`
- `output/<run-tag>_draft_toc_en.csv`
- `output/<run-tag>_draft_toc_en.txt`
- `output/<run-tag>_final_ordered.docx`

## Проверки после финальной сборки

- Основной текст не изменён.
- Добавленные разделы находятся только после статей.
- Нумерация страниц appendices сквозная.
- RU/EN TOC в одну колонку.
- Keyword formulas сохраняют hyphen и subscript.
- EN TOC контрольные pages: `117`, `180`, `196`, `223`, `232`, `312`, `419`.
- Перед коммитом обязательно открыть output docx вручную и проверить основной текст, нумерацию appendices и оглавления.

## Известные особенности

- Word COM может падать с ошибкой `Вызов был отклонен`, если Word занят; закрыть Word/убить `WINWORD` и повторить.
- `pywin32` `gen_py` cache может ломаться; очистить через `gencache.GetGeneratePath()`.
- Некоторые scripts из папки `scripts/` лучше запускать через `python -m scripts.<name>`, чтобы импорты работали корректно.
- Полная сборка может занимать несколько часов.
