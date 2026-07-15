# Supplier Price Intelligence Demo

[![Tests](https://github.com/Absamad-dew/price-intelligence-demo/actions/workflows/tests.yml/badge.svg)](https://github.com/Absamad-dew/price-intelligence-demo/actions/workflows/tests.yml)
[![Release](https://img.shields.io/github/v/release/Absamad-dew/price-intelligence-demo)](https://github.com/Absamad-dew/price-intelligence-demo/releases/tag/v0.1.1)

Воспроизводимый прототип под [живой заказ FL.ru](https://www.fl.ru/projects/5510010/razrabotka-bazyi-dannyih-i-sistemyi-sravneniya-praysov-postavschikov-python-ai.html): разные прайсы поставщиков приводятся к единому каталогу, пересчитываются к базовой единице, сравниваются по цене и сохраняются с полным audit trail.

В демо нет внешнего API и обязательных Python-зависимостей. CSV и плоские XLSX читаются и записываются стандартной библиотекой Python.

**Runnable proof на версионных демо-данных:** [system analysis case](docs/system-analysis-case.md) · [best-price CSV](output/best_prices.csv) · [review queue](output/review_queue.csv) · [quality metrics](output/quality_metrics.json) · [Excel-отчёт](output/price_intelligence.xlsx) · [release v0.1.1](https://github.com/Absamad-dew/price-intelligence-demo/releases/tag/v0.1.1). Это не заявление о production-качестве.

## Быстрый запуск

Требуется Python 3.10+.

```powershell
python price_demo.py
python -m unittest discover -s tests -v
```

Команда по умолчанию читает `samples`, использует `config.json` и пишет результаты в `output`. Для другого каталога:

```powershell
python price_demo.py --input-dir samples --output-dir output --config config.json
```

## Что доказано демо

1. Смешанный импорт: `supplier_a.csv` и настоящий `supplier_b.xlsx` обрабатываются одним запуском.
2. Разные схемы колонок приводятся к единому контракту через версионируемый mapping.
3. Единицы приводятся к базовым: `g → kg`, `ml → l`, `pcs/шт → pc`.
4. Rule-based fuzzy matching работает детерминированно и без LLM/API.
5. Цена пакета пересчитывается в цену за базовую единицу.
6. Низкая уверенность, несовпадение единиц и битые строки уходят в review queue.
7. Каждый импорт получает SHA-256 fingerprint от содержимого файла, mapping и его версии.
8. Повторный запуск с теми же данными сохраняет те же fingerprints и не дублирует строки.
9. Качество измеряется на вручную размеченном gold sample.

## Входные данные и mappings

- `samples/catalog.csv` — эталонный каталог;
- `samples/supplier_a.csv` — CSV-прайс;
- `samples/supplier_b.xlsx` — XLSX-прайс, лист `Offers`;
- `samples/supplier_b.csv` — тот же прайс в читаемом CSV-виде для аудита sample;
- `samples/gold_matches.csv` — ручная разметка ожидаемых matching/review решений;
- `mappings/v1.json` — mapping версии `1.0.0`.

Чтобы изменить схему поставщика, скопируйте `mappings/v1.json` в новый файл, увеличьте `mapping_version`, измените колонки и укажите новый `mapping_file` в `config.json`. Изменение версии или правил меняет fingerprint, даже если исходный прайс не менялся.

## Идемпотентный импорт

Fingerprint рассчитывается как:

```text
SHA256(file_bytes + canonical_supplier_mapping + mapping_version)
```

Он сохраняется в каждой нормализованной строке и в `output/import_manifest.json`. При повторном запуске `unchanged_imports` показывает количество уже известных импортов; итоговые строки и best-price выбор не размножаются.

## Метрики качества

`samples/gold_matches.csv` содержит шесть известных товаров и две строки, которые должны попасть на ручную проверку.

- `precision = correct_auto_matches / auto_matched_rows`;
- `coverage = correct_auto_matches / matchable_gold_rows`;
- `review_share = reviewed_or_rejected_rows / gold_rows`.

Текущий sample даёт:

| Метрика | Значение | Gate |
|---|---:|---:|
| Precision | 100% | ≥ 95% |
| Coverage | 100% | ≥ 90% |
| Review share | 25% | ≤ 35% |

Это не заявление о production-качестве: перед пилотом gold sample нужно расширить до 50–100 обезличенных реальных строк заказчика.

## Результаты

- `output/normalized_offers.csv` — все нормализованные предложения;
- `output/best_prices.csv` — лучший поставщик по каждому товару;
- `output/review_queue.csv` — ручная проверка и причины;
- `output/quality_metrics.json` — precision, coverage, review share и quality gate;
- `output/import_manifest.json` — версии mapping и fingerprints;
- `output/run_summary.json` — счётчики запуска;
- `output/price_intelligence.xlsx` — те же результаты в одном Excel-файле на пяти листах.

## Границы прототипа

Встроенный XLSX-адаптер предназначен для обычных плоских таблиц: первая строка — заголовки, далее данные. Для сложных Excel-файлов с формулами, объединениями и несколькими служебными секциями production-версия должна использовать отдельный адаптер на `openpyxl`/`pandas` и сохранять исходный файл неизменным.

Следующие этапы production MVP: PostgreSQL, API загрузки, права пользователей, журнал исправлений, acceptance dataset и embeddings только для строк ниже rule-based порога.

Контакты: Telegram `@Absamad_m`, Gmail `absamad.manturov@gmail.com`, GitHub `Absamad-dew`.

Лицензия: MIT.
