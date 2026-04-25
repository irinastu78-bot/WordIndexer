# Полный прогон pipeline для нового Word-файла

Эта инструкция описывает полный ручной прогон WordKeywords для нового сборника статей из PowerShell. Пример ниже использует файл `input\test7.docx` и run-tag `test7`.

## Требования к исходному Word-файлу

- Файл должен быть в формате `.docx`.
- Это должен быть один общий Word-документ со всеми статьями.
- В каждой статье должны быть русскоязычные и англоязычные блоки.
- RU keywords должны начинаться с `Ключевые слова:`.
- EN keywords должны начинаться с `Keywords:`, `Key words:` или другого поддерживаемого варианта, например `Ketwords:`.
- Заголовки и авторы должны находиться рядом с keyword-блоками примерно как в текущем поддерживаемом шаблоне.
- Во время записи output-файлов итоговый `.docx` не должен быть открыт в Microsoft Word.
- Новый файл нужно положить в `input\`, например `input\test7.docx`.
- Если файл скопирован неправильно и имеет размер `0`, pipeline не сможет корректно обработать документ.

## Полный порядок команд

Запускать из корня репозитория:

```powershell
python dump_doc_paragraph_snapshot.py --docx input\test7.docx --run-tag test7
```

Что делает: создаёт снимок параграфов Word-документа с текстом, страницами, стилями и диагностикой.

Главный output: `output\test7_doc_paragraph_snapshot.csv`.

Нормальный summary: snapshot создан, количество строк соответствует размеру документа.

Если ошибка: проверить путь `input\test7.docx`, размер файла и доступность Microsoft Word.

```powershell
python build_author_title_paragraph_ru_from_snapshot.py --run-tag test7
```

Что делает: находит RU статьи, заголовки и авторские строки по snapshot.

Главный output: `output\test7_author_title_paragraph_ru_debug_from_snapshot.csv`.

Нормальный summary: `articles found` равно ожидаемому числу статей, например `95`.

Если ошибка: проверить, что существует `output\test7_doc_paragraph_snapshot.csv` и что run-tag указан одинаково.

```powershell
python -m scripts.debug_author_windows_en --run-tag test7
```

Что делает: строит локальные EN search windows на основе найденных RU статей.

Главный output: `output\test7_author_windows_en_debug.csv`.

Нормальный summary: `WINDOWS DEFINED` равно ожидаемому числу статей.

Если ошибка: проверить RU debug CSV из предыдущего шага.

```powershell
python -m scripts.debug_en_title_author_pairs --docx input\test7.docx --run-tag test7
```

Что делает: ищет EN title-author pairs внутри локальных окон.

Главный output: `output\test7_en_title_author_pairs_debug.csv`.

Нормальный summary: `TITLE AND AUTHOR FOUND` равно ожидаемому числу статей.

Если ошибка: смотреть `output\test7_en_title_author_pairs_debug.txt` и строки со статусом `not_found`.

```powershell
python -m scripts.build_draft_author_index_en --docx input\test7.docx --run-tag test7
```

Что делает: строит черновой EN author index из EN title-author pairs.

Главный output: `output\test7_draft_author_index_en.txt`.

Нормальный summary: `EN ARTICLES PROCESSED` равно числу статей, `NOT FOUND: 0`.

Если ошибка: смотреть `output\test7_draft_author_index_en_debug.txt` и проверять, почему авторская строка не распарсилась.

```powershell
python -m scripts.enrich_ru_title_paragraph_structure --docx input\test7.docx --run-tag test7
```

Что делает: добавляет структуру RU title paragraph, включая run fragments и признаки заголовок+авторы.

Главный output: `output\test7_ru_title_paragraph_structure_debug.csv`.

Нормальный summary: файл создан, строки соответствуют найденным RU статьям.

Если ошибка: проверить snapshot и RU title-author debug artifacts.

```powershell
python build_draft_author_index_ru_from_snapshot.py --run-tag test7
```

Что делает: строит черновой RU author index по snapshot и найденным RU статьям.

Главный output: `output\test7_draft_author_index_ru_debug_from_snapshot.txt`.

Нормальный summary: `RU ARTICLES PROCESSED` равно числу статей, `NOT FOUND: 0`.

Если ошибка: смотреть debug TXT/CSV и проверять проблемные author blocks.

```powershell
python build_author_index_ru_text_from_snapshot.py --run-tag test7
```

Что делает: формирует финальный текст RU author index для вставки в Word.

Главный output: `output\test7_author_index_ru_from_snapshot.txt`.

Нормальный summary: файл создан, авторы сгруппированы со страницами.

Если ошибка: проверить наличие чернового RU author debug output из предыдущего шага.

```powershell
python debug_toc_ru_from_word.py --docx input\test7.docx --run-tag test7
```

Что делает: строит draft RU TOC из Word-документа.

Главный output: `output\test7_draft_toc_ru.csv`.

Нормальный summary: `rows read` или найденные строки TOC соответствуют ожидаемому числу статей.

Если ошибка: проверить, что Word доступен, файл `.docx` не повреждён, а заголовки похожи на поддерживаемый шаблон.

```powershell
python build_toc_ru_draft_text.py --run-tag test7
```

Что делает: создаёт plain-text черновик RU TOC для диагностики.

Главный output: `output\test7_draft_toc_ru.txt`.

Нормальный summary: `rows read` равно ожидаемому числу строк TOC.

Если ошибка: проверить наличие `output\test7_draft_toc_ru.csv`.

```powershell
python fast_keyword_index_find.py --docx input\test7.docx --run-tag test7
```

Что делает: через Word COM ищет RU/EN keyword paragraphs и страницы.

Главный output: `output\test7_fast_keyword_rows_v2.csv`.

Нормальный summary: RU/EN hits примерно равны числу статей; `HITS WITHOUT PAGE` должен быть `0` или объяснимым.

Если ошибка: проверить keyword labels в документе и убедиться, что `--docx` указывает на нужный файл.

```powershell
python build_separate_keyword_indexes.py --run-tag test7
```

Что делает: строит отдельные RU и EN keyword indexes из fast keyword rows.

Главный output: `output\test7_keyword_index_ru.txt` и `output\test7_keyword_index_en.txt`.

Нормальный summary: есть ненулевое количество RU и EN unique keywords.

Если ошибка: проверить наличие `output\test7_fast_keyword_rows_v2.csv`.

```powershell
python insert_keyword_indexes_into_word.py --docx input\test7.docx --run-tag test7
```

Что делает: вставляет RU и EN keyword indexes в копию исходного Word-документа.

Главный output: `output\test7_with_keyword_indexes.docx`.

Нормальный summary: создан `.docx`; если основной файл занят, может быть создан fallback с суффиксом `_1`, `_2` и так далее.

Если ошибка: закрыть Word-файлы, проверить наличие `output\test7_keyword_index_ru.txt` и `output\test7_keyword_index_en.txt`.

```powershell
python insert_toc_and_author_into_word.py --docx output\test7_with_keyword_indexes.docx --run-tag test7
```

Что делает: вставляет RU TOC и RU author index в документ с keyword indexes.

Главный output: `output\test7_final_with_toc_author.docx`.

Нормальный summary: TOC entries и author lines вставлены, итоговый файл сохранён.

Если ошибка: закрыть итоговый `.docx`, проверить fallback output и наличие `output\test7_author_index_ru_from_snapshot.txt`.

```powershell
python insert_en_blocks_into_word.py --docx output\test7_final_with_toc_author.docx --run-tag test7
```

Что делает: вставляет EN author index в финальный документ.

Главный output: `output\test7_final_with_toc_author_en.docx`.

Нормальный summary: EN author index inserted, итоговый `.docx` сохранён.

Если ошибка: проверить наличие `output\test7_draft_author_index_en.txt` и закрыть открытые Word-файлы.

## Контрольные признаки успешного прогона

- `articles found: 95` или другое ожидаемое число статей для конкретного сборника.
- `WINDOWS DEFINED` равно числу статей.
- `TITLE AND AUTHOR FOUND` равно числу статей.
- `EN ARTICLES PROCESSED` равно числу статей.
- `EN NOT FOUND: 0` или `NOT FOUND: 0` в EN author step.
- `RU NOT FOUND: 0` в RU author diagnostics.
- RU keyword hits и EN keyword hits равны числу статей или объяснимо близки к нему.
- Итоговый файл создан: `output\test7_final_with_toc_author_en.docx`.
- В итоговом файле есть RU keyword index, EN keyword index, RU TOC, RU author index и EN author index.

## Типовые ошибки

- `FileNotFoundError` по промежуточному CSV обычно означает, что пропущен предыдущий шаг или указан неверный `--run-tag`.
- Если `--help` запускает обработку, значит у скрипта старый CLI и его нужно исправлять до запуска на рабочем документе.
- Если Word не сохраняет файл, закрыть открытый `.docx` и проверить, появился ли fallback output с суффиксом `_1`, `_2` и так далее.
- Если `NOT FOUND > 0`, смотреть соответствующие debug `.txt` и `.csv`.
- Если input `.docx` имеет размер `0`, файл скопирован неправильно.
- Если скрипт ищет не те tagged artifacts, проверить единый `--run-tag` во всех командах.
- Если число статей неожиданно меньше ожидаемого, начать с `test7_doc_paragraph_snapshot.csv` и RU/EN debug artifacts.

## Что не коммитить

- `output\*.docx`.
- Временные Word-файлы, включая файлы вида `~$*.docx`.
- Большие промежуточные артефакты из `output\`, если они уже покрыты `.gitignore`.
- Локальные ручные копии итоговых документов.
