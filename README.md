# WordKeywords

Минимально актуальная структура проекта:

- `fast_keyword_index_find.py` - совместимый корневой запуск поиска keyword-блоков через Word COM `Find`
- `build_separate_keyword_indexes.py` - совместимый корневой запуск сборки RU/EN указателей
- `insert_keyword_indexes_into_word.py` - совместимый корневой запуск вставки объединенного указателя в Word
- `scripts/` - актуальная реализация рабочих скриптов
- `src/wordkeywords/common.py` - общие утилиты нормализации, CSV и базовой обработки keyword-строк
- `archive/` - старые промежуточные и экспериментальные версии, не входящие в текущий рабочий пайплайн

Актуальный пайплайн:

1. `python fast_keyword_index_find.py`
2. `python build_separate_keyword_indexes.py`
3. `python insert_keyword_indexes_into_word.py`

Текущее состояние `test4`:

- входной файл заказчика: `input/test4.docx`
- snapshot pipeline для `test4` собран
- RU author draft from snapshot собран
- RU TOC draft собран
- keyword pipeline собран
- финальный документ собирается
- RU TOC author wrapping/layout доведён до приемлемого состояния
- в финальной сборке восстановлены предметные указатели

Порядок итоговых артефактов для `test4`:

1. `output/test4_with_keyword_indexes.docx`
2. `output/test4_final_with_toc_author.docx`

Примечание по финальной сборке:

- финальная вставка TOC/author в `insert_toc_and_author_into_word.py` использует `output/<run-tag>_with_keyword_indexes.docx` как source document, если такой файл уже существует для текущего `run-tag`
- если такого tagged файла нет, используется явно переданный `--docx`

Следующий приоритет:

1. `EN author index`

Архивные файлы сохранены без удаления:

- `archive/build_keyword_index.py` - старая версия извлечения через перебор абзацев
- `archive/export_keywords_csv.py` - ранний экспорт через `python-docx`
- `archive/reconcile_with_toc.py` - промежуточный скрипт сверки с оглавлением
- `archive/tt.py` - отладочный/диагностический скрипт
