FE-RUS SEO Manager v1.0

Один GUI вместо разрозненных скриптов.

Этапы:
1. Категория - URL -> реальный ocFilterData -> option/value/value_id/params.
2. WordKeeper - загрузка TOP-30 CSV.
3. Анализ - сопоставление кандидатов с WordKeeper и получение competitor_1..5.
4. OCFilter - подготовлен отдельный этап. В v1.0 создание намеренно заблокировано.
5. Экспорт - единый XLSX.

Важно:
- URL target_url не выдумывается.
- Наличие значения OCFilter не означает CREATE.
- CREATE_CANDIDATE появляется при наличии конкурентных URL в WordKeeper.
- EXISTS должен подтверждаться живым OCFilter/OpenCart на следующем этапе.
- Настройки и пароли сохраняются в %APPDATA%\FE-RUS SEO Manager\config.json.
- Для паролей Ctrl+V обработан вручную.
