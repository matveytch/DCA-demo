# Как мы проверяем проект

Что считается рабочим кодом и как Claude проверяет себя сам. Актуально на 2026-11-09.

## Цель

Защитить главный сценарий: оператор заполнил Excel → получил точный пакет документов, сделка лежит в папке компании, реестр партий верен.
Ошибка в сумме, количестве или реквизитах стоит денег. Поэтому каждое число и каждый путь проверяет тест.

## Как запустить тесты

```bash
python3 -m pip install -r requirements-lock.txt
python3 -m pytest
```

- Тесты идут параллельно, по процессу на ядро (`pytest-xdist`). Полный прогон — меньше минуты.
- Последняя строка должна быть вида «353 passed». Слова `failed` или `error` значат, что-то сломано.
- Проверено 2026-11-09: «353 passed».
- Тест с настоящим restic пропускается, если restic не установлен.

Повторная проверка, актуально на 2026-11-09:

- Создано чистое окружение Python 3.13.9 без системных пакетов.
- Установлены зависимости из `requirements-lock.txt` и собранный пакет `dca==0.1.0`.
- `pip check` не нашёл несовместимых зависимостей.
- Полный набор: `353 passed`, включая проверку настоящего restic.
- Пример из README создал документы, ZIP, журнал и две записи реестра.
- Проверены суммы: 450 USD = 270 USD + 180 USD.
- Независимая проверка обезличивания завершена. Gitleaks 8.30.1 не нашёл секретов в проверенных файлах.
- Дополнительно проверено распакованное содержимое всех семи Office-файлов.

Другие варианты:

| Задача | Команда |
|---|---|
| Один файл тестов | `python3 -m pytest tests/test_render.py` |
| Один тест | `python3 -m pytest tests/test_render.py::test_dates_declension_and_amount_in_words` |
| Тесты по слову в имени | `python3 -m pytest -k ledger` |
| По одному, для отладки | `python3 -m pytest -n 0 tests/test_cli.py` |
| Время каждого теста | `python3 -m pytest --durations=20` |

Шаблоны и Excel-файлы создаёт скрипт. После правки скрипта запустите его и тесты:

```bash
python3 scripts/make_demo_templates.py
python3 -m pytest
```

## Где тесты запускаются автоматически

- CI на GitHub: `.github/workflows/tests.yml`, на каждый push и pull request. Версии — из `requirements-lock.txt`.
- Перед каждым коммитом Claude запускает полный набор (правило из `CLAUDE.md`).

## Обязательные правила

1. **Баг закрывается только вместе с тестом.**
   1. Воспроизведите проблему.
   2. Зафиксируйте её тестом.
   3. Исправьте код.
   4. Убедитесь: тест падал на старом поведении и проходит на новом.
2. **Новое требование — сначала тест.** Запишите требование в `docs/requirements.md`, добавьте тест, затем код.
3. **Не подгоняйте тест под код.** Тест, изменённый ради прохождения, прячет поломку. Исправляйте код, а не проверку. Меняйте тест, только если изменилось требование и это записано в `docs/requirements.md`.
4. **Коммит — только при зелёных тестах.** Читайте итоговую строку pytest, а не код возврата конвейера команд.
5. **Тесты не трогают реальные данные.** Только временные папки (`tmp_path`, `DCA_HOME`) и выдуманные данные из `data/input_example_FILLED.xlsx`.
6. **Галочку у критерия приёмки** в `docs/requirements.md` ставьте только после проверки по карте ниже.

## Из чего состоит набор

| Файл | Что проверяет |
|---|---|
| `tests/helpers.py` | Общие помощники: пример сделки, правка Excel, текст docx |
| `test_models.py` | Модель сделки: значения по умолчанию, бизнес-правила |
| `test_excel_io.py` | Чтение Excel: справочники, партии, приложения, листы-формы, сообщения об ошибках |
| `test_render.py` | Демо-шаблоны: условные абзацы, строки таблиц, даты, склонение, прописью, строгие имена |
| `test_invariants.py` | Итоги на сетке цен и количеств; один файл на приложение |
| `test_cli.py` | Команды от начала до конца: new, generate, amend, transfer, from-log, runner'ы, doctor |
| `test_attachments.py` | Пакет приложений и мягкие проверки |
| `test_ledger*.py` | Реестр партий, «Баланс», `check-ledger` |
| `test_master_db.py` | База справочников: создание, снимок в сделку, `refresh-db` |
| `test_number_change.py` | Смена номера соглашения, журнал запусков |
| `test_output_visible.py` | Папка сборки без точки, флаг «скрытый» |
| `test_backup.py` | Бэкап и восстановление (restic подменён заглушкой; один тест — с настоящим) |
| `test_paths.py` | Где код и где данные; запуск без установки после переноса |
| `test_archive_log.py` | Архив прежних версий, `log.json`, ZIP |
| `test_textfit.py` | Раскладка текста по строкам граф фиксированной ширины |
| `test_claude_hook.py` | Хук, который держит сессию Claude внутри репозитория |
| `test_smoke.py` | Загрузка пакета `dca`, даты, склонение, сумма прописью |

## Карта: критерий приёмки → тест

| Критерий | Тесты |
|---|---|
| AC1 | `test_cli.py::test_new_creates_folder_identity_input_and_portable_runners`, `::test_new_refuses_an_existing_folder_and_a_duplicate_number`, `test_master_db.py::test_new_copies_all_records_and_deal_loads` |
| AC2 | `test_cli.py::test_generate_full_pipeline`, `test_render.py::test_outgoing_package_files` |
| AC3 | `test_excel_io.py::test_all_problems_in_one_message_with_sheet_and_row`, `test_cli.py::test_invalid_input_exits_1_with_one_russian_message_and_no_documents` |
| AC4 | `test_invariants.py::test_total_is_printed_price_times_quantity_rounded_up`, `test_render.py::test_dates_declension_and_amount_in_words` |
| AC5 | `test_models.py::test_only_kg_and_usd` |
| AC6 | `test_cli.py::test_second_generate_archives_the_previous_output`, `::test_failure_while_building_leaves_the_previous_output`, `::test_failed_publish_restores_the_previous_output` |
| AC7 | `test_models.py::test_annex_quantities_must_sum_and_numbers_be_unique`, `test_invariants.py::test_one_annex_file_per_annex`, `::test_multi_annex_total_is_the_sum_of_annex_totals` |
| AC8 | `test_attachments.py::test_copy_renames_by_number_and_label`, `test_render.py::test_cover_list_numbering_matches_the_batches`, `test_cli.py::test_generate_full_pipeline` |
| AC9 | `test_attachments.py::test_batch_of_another_company_and_without_owner`, `::test_incomplete_batch_and_no_batch_at_all`, `::test_complete_deal_has_no_notices`, `test_cli.py::test_missing_scans_are_notices_in_the_check_file` |
| AC10 | `test_ledger.py::test_append_creates_new_ledger`, `::test_balance_sheet_has_formulas`, `test_ledger_check.py::test_warnings_double_write_off_different_weight_overuse` |
| AC11 | `test_ledger_balance.py::test_groups_by_company_used_first_then_unused`, `::test_sheet_has_company_headings_formulas_and_light_green_unused` |
| AC12 | `test_ledger_check.py` (все тесты) |
| AC13 | `test_cli.py::test_transfer_writes_the_act_only`, `test_render.py::test_transfer_needs_a_filled_sheet` |
| AC14 | `test_cli.py::test_unfilled_transfer_sheet_is_not_a_notice`, `::test_filled_transfer_sheet_with_an_error_is_a_notice` |
| AC15 | `test_render.py::test_transfer_text_that_does_not_fit_is_refused_with_every_problem`, `test_textfit.py` |
| AC16 | `test_cli.py::test_generate_incoming_pipeline`, `::test_new_incoming_hides_outgoing_rows_and_the_transfer_sheet`, `::test_the_command_must_match_the_deal_kind` |
| AC17 | `test_cli.py::test_amend_writes_annexes_and_the_amendment`, `::test_amend_refuses_an_incomplete_sheet_and_keeps_the_previous_result`, `test_render.py::test_amendment_numbers_only_the_clauses_that_change` |
| AC18 | `test_cli.py::test_from_log_re_renders_and_refuses_a_newer_schema`, `::test_from_log_reports_changed_templates` |
| AC19 | `test_cli.py::test_stale_signatures_are_reported` |
| AC20 | `test_ledger_check.py::test_cli_refuses_while_excel_has_the_ledger_open`, `test_master_db.py::test_refresh_db_refuses_open_excel_and_missing_db`, `test_cli.py::test_excel_open_keeps_the_deal_where_it_is` |
| AC21 | `test_cli.py::test_runner_works_in_a_moved_copy_without_dca_on_path`, `test_paths.py::test_launcher_runs_the_code_next_to_it_after_a_move` |
| AC22 | `test_backup.py` (все тесты) |
| AC23 | `test_render.py::test_variant_b_adds_the_party_c_clause_and_nothing_else_changes`, `test_models.py::test_variant_b_needs_all_of_party_c` |
| AC24 | `test_number_change.py` |
| AC25 | `test_number_change.py::test_missing_output_after_filing_is_a_red_error_not_done`, `::test_every_run_writes_a_line_to_the_generate_log`, `test_output_visible.py` |
| AC26 | `test_attachments.py::test_date_contradictions_are_warnings`, `test_cli.py::test_transfer_date_after_due_date_is_a_warning` |
| AC27 | `test_render.py::test_unknown_template_variable_is_an_error`, `::test_nested_typo_is_an_error_too` |
