# Рефакторинг Loader1C — План

## Текущее состояние
- Один файл `Program.cs` ~2000+ строк
- Вся логика смешана: CLI, COM-подключение, парсинг файлов, создание документов, справочники

## Предлагаемая структура

### 1. `Program.cs` (~100 строк)
- Точка входа, разбор CLI-команд (load, scan, inspect, meta, banks, fix-bankacct)
- Настройка Host для Windows Service
- Делегирование логики в соответствующие классы

### 2. `LoaderConfig.cs` (~30 строк)
- Класс конфигурации (уже есть, просто вынести)
- `AccountMapping` класс

### 3. `BankFileParser.cs` (~200 строк)
- Парсинг 1CClientBankExchange TXT формата
- `BankExchangeFile`, `BankDocument` классы

### 4. `Com1CConnector.cs` (~800 строк)
- Подключение к базам 1С через COM
- `ScanDatabases()`, `LoadFile()`, маппинг счетов
- `FindAccount()`, `FindBankAccount()`, `FindBankAccountByNumber()`
- `CountExistingDocs()`
- `NormalizeAccount()`

### 5. `DocumentCreator.cs` (~500 строк)
- `CreateDocument()` — основная логика создания СписаниеСРасчетногоСчета / ПоступлениеНаРасчетныйСчет
- Определение внутреннего перевода
- Поиск/создание контрагентов и банковских счетов (`EnsureCounterparties`)
- Установка ВидОперации, СтатьяДДС, СчетУчета и т.д.
- Работа с зарплатными ведомостями

### 6. `LoaderWorker.cs` (~100 строк)
- `BackgroundService` для Windows Service режима
- Цикл сканирования `data/output/`
- Отслеживание обработанных файлов

### 7. `FileLoggerProvider.cs` (~50 строк)
- Простой файловый логгер (уже есть, вынести)

## Приоритеты
1. Вынести `BankFileParser` и модели данных — чистая логика, легко выделить
2. Вынести `LoaderWorker` — уже логически отделён
3. Разделить `Com1CConnector` и `DocumentCreator` — основная работа
4. CLI-команды оставить в `Program.cs` тонкими обёртками

## Риски
- COM interop через dynamic — сложно тестировать
- Нужно аккуратно с shared state (_currentPayrollRef, _accountMap)
- После рефакторинга обязательно протестировать load на реальной выписке
