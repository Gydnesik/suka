# process-schedule

Edge Function принимает от администратора таблицу расписания `.xlsx`, `.ods` или `.csv`.

Пайплайн:
1. браузер отправляет файл в Edge Function только после авторизации администратора;
2. Edge Function читает таблицу через лёгкий ZIP/XML-парсер (`fflate`) без SheetJS;
3. выбирается наиболее похожий на расписание лист;
4. объединённые ячейки раскрываются в двумерную сетку;
5. сетка конвертируется в чистый CSV;
6. только CSV-текст отправляется в Gemini;
7. нормализованный JSON передаётся в `replace_schedules()`;
8. RPC в одной транзакции удаляет ВСЕ старые строки `schedules` и вставляет только новое расписание.

Gemini получает строгие инструкции игнорировать дежурство, ВПР, примечания и прочую мета-информацию, искать нижнюю основную сетку, читать классы 5А–11Б вертикально, сохранять номер урока и время, сохранять `-` и корректно обрабатывать объединённые ячейки.

## Secrets

В Supabase:

```bash
supabase secrets set GEMINI_API_KEY="YOUR_GEMINI_KEY" GEMINI_MODEL="gemini-3.6-flash"
```

`GEMINI_API_KEY` хранится только на стороне Edge Function.

## Deploy

Перед первым деплоем новой версии обязательно выполни обновлённый `schema.sql` в Supabase SQL Editor — он создаёт RPC `replace_schedules(jsonb)`.

```bash
supabase link --project-ref igbkjkjagkhxpxezjwtj
supabase functions deploy process-schedule --project-ref igbkjkjagkhxpxezjwtj
```


## v6.2.0 Memory fix

В этой версии полностью удалён импорт SheetJS. Он мог сам по себе давать высокий пик памяти Edge Function. ODS и XLSX читаются напрямую как ZIP/XML через `fflate`; CSV читается потоково. Обрабатывается только безопасная область до 2500 строк × 80 колонок. Формат `.xls` намеренно не принимается: сохрани его как `.xlsx` или `.ods`.
