// FiCalc — независимый расчётчик черногорской финансовой отчётности из данных 1С.
//
// Назначение: кросс-проверка внешней обработки FinIskaziCG (1С:Бухгалтерия КОРП МСФО).
// Два независимых движка (эта утилита и обработка 1С) считают одни и те же строки отчётности
// по одному и тому же mapping.json и одному и тому же регистру бухгалтерии — цифры должны
// совпасть. FiCalc НЕ читает и не воспроизводит логику обработки 1С, а строит результат
// напрямую из проводок РегистрБухгалтерии.МСФО.
//
// Использование:
//   FiCalc.exe <db> <orgName> <year> <mappingPath> <outJsonPath>
// Пример:
//   FiCalc.exe base_m "MINDER DOO" 2025 C:\scripts\fi-epf\mapping.json C:\scripts\fi-epf\out\minder_check.json
//
// Алгоритм:
//   1. Читает mapping.json (структура секций/строк описана в SPEC.md и в reference/mapping.json).
//   2. Подключается к 1С через V83.COMConnector и выполняет 4 запроса к
//      РегистрБухгалтерии.МСФО: Остатки на конец Года и Года-1, Обороты за Год и за Год-1.
//   3. Для каждой account-строки агрегирует остатки/обороты счетов, чей код начинается с
//      любого include-префикса и не начинается ни с одного exclude-префикса, по типу balance.
//   4. Формульные строки считает после всех account-строк, итеративно (формулы могут
//      ссылаться на другие формулы) — до неподвижной точки.
//   5. Пишет результат в JSON (UTF-8) и печатает в консоль контрольные строки.
//
// Осознанные допущения (см. также отчёт задачи):
//   - Обрабатываются и выгружаются ТОЛЬКО секции BilanStanja, BilanUspjeha, StatistickiAneks —
//     это единственные секции mapping.json, где вообще встречаются строки kind="account"
//     (остальные секции в v1 обработки заполняются вручную/нулями — см. SPEC.md, "Ограничения
//     v1" — кросс-проверка для них не имеет смысла, т.к. оба движка дадут одни нули).
//   - Поле accounts.dio ("счёт входит частично, разносится корректировками в живой обработке")
//     НЕ влияет на агрегацию FiCalc — используются только include/exclude, как явно указано в
//     задании. Это осознанно воспроизводит поведение v1 обработки 1С (вся сумма dio-счёта
//     попадает в первичную строку), поэтому кросс-проверка остаётся корректной.
//   - Защита от задвоения счёт-группа/субсчёт: после каждой выборки остатков/оборотов в C#
//     остаются только "листовые" коды — коды, не являющиеся префиксом другого, более длинного
//     кода из той же выборки (см. LeafOnly). Правило применяется отдельно к каждой из 4 выборок.
//   - Семантика 3-й колонки (Iznos3) секции BilanStanja в ТЗ не определена — она всегда 0;
//     реальная семантика колонок документирована в выходном JSON в поле "_meta".

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace FiCalcTool
{
    #region Модель mapping.json

    internal sealed class MappingRoot
    {
        [JsonPropertyName("sections")]
        public Dictionary<string, MappingSection> Sections { get; set; } = new();
    }

    internal sealed class MappingSection
    {
        [JsonPropertyName("rows")]
        public List<MappingRow> Rows { get; set; } = new();
    }

    internal sealed class MappingRow
    {
        [JsonPropertyName("redniBroj")]
        public string RedniBroj { get; set; } = "";

        [JsonPropertyName("pozicija")]
        public string Pozicija { get; set; } = "";

        [JsonPropertyName("grupaRacuna")]
        public string GrupaRacuna { get; set; } = "";

        [JsonPropertyName("kind")]
        public string Kind { get; set; } = "";

        [JsonPropertyName("formula")]
        public string? Formula { get; set; }

        [JsonPropertyName("accounts")]
        public AccountsSpec? Accounts { get; set; }

        [JsonPropertyName("balance")]
        public string? Balance { get; set; }
    }

    internal sealed class AccountsSpec
    {
        [JsonPropertyName("include")]
        public List<string>? Include { get; set; }

        [JsonPropertyName("exclude")]
        public List<string>? Exclude { get; set; }

        // Не используется в расчёте FiCalc — см. комментарий в шапке файла.
        [JsonPropertyName("dio")]
        public List<string>? Dio { get; set; }
    }

    #endregion

    #region Факты из регистра

    /// <summary>
    /// Дт/Кт остатка на дату (Остатки) либо Дт/Кт оборота за период (Обороты) для одного счёта.
    /// Обе выборки в запросах называют колонки одинаково (Код/Дт/Кт), поэтому используем одну
    /// структуру и один код чтения для обоих случаев.
    /// </summary>
    internal readonly struct AccountAmounts
    {
        public readonly string Code;
        public readonly decimal Debit;
        public readonly decimal Credit;

        public AccountAmounts(string code, decimal debit, decimal credit)
        {
            Code = code;
            Debit = debit;
            Credit = credit;
        }
    }

    /// <summary>Четыре факт-набора, из которых считаются все account-строки всех секций.</summary>
    internal sealed class SectionFacts
    {
        private readonly Dictionary<string, AccountAmounts> _remaindersYear;
        private readonly Dictionary<string, AccountAmounts> _remaindersPrevYear;
        private readonly Dictionary<string, AccountAmounts> _turnoversYear;
        private readonly Dictionary<string, AccountAmounts> _turnoversPrevYear;

        public SectionFacts(
            Dictionary<string, AccountAmounts> remaindersYear,
            Dictionary<string, AccountAmounts> remaindersPrevYear,
            Dictionary<string, AccountAmounts> turnoversYear,
            Dictionary<string, AccountAmounts> turnoversPrevYear)
        {
            _remaindersYear = remaindersYear;
            _remaindersPrevYear = remaindersPrevYear;
            _turnoversYear = turnoversYear;
            _turnoversPrevYear = turnoversPrevYear;
        }

        /// <summary>Выбирает нужный факт-набор по типу balance строки и колонке (текущий/предыдущий год).</summary>
        public Dictionary<string, AccountAmounts> Select(string balanceKind, bool previousYear)
        {
            bool isTurnover = balanceKind == "turnoverDebitNet" || balanceKind == "turnoverCreditNet";
            if (isTurnover) return previousYear ? _turnoversPrevYear : _turnoversYear;
            return previousYear ? _remaindersPrevYear : _remaindersYear;
        }
    }

    #endregion

    /// <summary>Один слагаемый формулы: знак + ссылка на redniBroj другой строки той же секции.</summary>
    internal readonly struct FormulaTerm
    {
        public readonly bool Negative;
        public readonly string Reference;

        public FormulaTerm(bool negative, string reference)
        {
            Negative = negative;
            Reference = reference;
        }
    }

    #region Выходной JSON

    internal sealed class FiCalcOutput
    {
        [JsonPropertyName("org")]
        public string Org { get; set; } = "";

        [JsonPropertyName("year")]
        public int Year { get; set; }

        [JsonPropertyName("sections")]
        public Dictionary<string, Dictionary<string, decimal[]>> Sections { get; set; } = new();

        [JsonPropertyName("_meta")]
        public FiCalcMeta Meta { get; set; } = new();
    }

    internal sealed class FiCalcMeta
    {
        // Семантика колонок Iznos1..IznosN по секциям (см. шапку файла и SPEC.md).
        [JsonPropertyName("columns")]
        public Dictionary<string, string> Columns { get; set; } = new();
    }

    #endregion

    internal static class Program
    {
        // Секции mapping.json, которые FiCalc реально считает и выгружает — единственные,
        // где встречаются строки kind="account" (см. шапку файла).
        private static readonly string[] ComputedSections = { "BilanStanja", "BilanUspjeha", "StatistickiAneks" };

        private static readonly List<string> EmptyStrings = new();

        private static readonly Regex FormulaTokenRegex = new(@"[+-]?[0-9]+[A-Za-z]*", RegexOptions.Compiled);

        private static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            if (args.Length < 5)
            {
                Console.Error.WriteLine("Использование: FiCalc.exe <db> <orgName> <year> <mappingPath> <outJsonPath>");
                return 1;
            }

            string db = args[0];
            string orgName = args[1];
            string yearArg = args[2];
            string mappingPath = args[3];
            string outJsonPath = args[4];

            if (!int.TryParse(yearArg, NumberStyles.Integer, CultureInfo.InvariantCulture, out int year))
            {
                Console.Error.WriteLine($"Некорректный аргумент year: '{yearArg}'.");
                return 1;
            }

            MappingRoot mapping;
            try
            {
                mapping = LoadMapping(mappingPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Не удалось прочитать mapping-файл '{mappingPath}': {ex.Message}");
                return 1;
            }

            foreach (string sectionName in ComputedSections)
            {
                if (!mapping.Sections.ContainsKey(sectionName))
                {
                    Console.Error.WriteLine($"В mapping-файле нет секции '{sectionName}', ожидаемой FiCalc.");
                    return 1;
                }
            }

            // Валидируем структуру mapping.json ДО подключения к 1С — незачем тратить время на
            // COM-сессию и запросы, если сам маппинг уже сломан (например строка kind="account"
            // с пустым accounts.include — такое реально встречалось в mapping.json). Собираем
            // ВСЕ найденные проблемы сразу, а не падаем на первой же строке.
            List<string> mappingProblems = ValidateMapping(mapping);
            if (mappingProblems.Count > 0)
            {
                Console.Error.WriteLine($"mapping.json не прошёл проверку — найдено проблем: {mappingProblems.Count}. Подключение к 1С не выполняется.");
                foreach (string problem in mappingProblems) Console.Error.WriteLine("  - " + problem);
                return 1;
            }

            object connObj;
            try
            {
                Type connectorType = Type.GetTypeFromProgID("V83.COMConnector")
                    ?? throw new InvalidOperationException("ProgID V83.COMConnector не зарегистрирован на этой машине.");
                dynamic connector = Activator.CreateInstance(connectorType)!;
                dynamic conn = connector.Connect(
                    $"Srvr=\"navus-server\";Ref=\"{db}\";Usr=\"Zaykov Andrey\";Pwd=\"19700214\"");
                connObj = (object)conn;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Не удалось подключиться к базе 1С '{db}': {ex.Message}");
                return 1;
            }

            try
            {
                return Run(connObj, db, orgName, year, mapping, outJsonPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("ОШИБКА:");
                Console.Error.WriteLine(ex.ToString());
                return 1;
            }
            finally
            {
                try { Marshal.FinalReleaseComObject(connObj); }
                catch { /* при завершении процесса это лучшее, что можно сделать */ }
            }
        }

        private static MappingRoot LoadMapping(string path)
        {
            string text = File.ReadAllText(path, Encoding.UTF8);
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            MappingRoot? root = JsonSerializer.Deserialize<MappingRoot>(text, options);
            if (root is null || root.Sections.Count == 0)
                throw new InvalidOperationException("mapping.json не содержит секций (пустой или некорректный формат).");
            return root;
        }

        /// <summary>
        /// Прогоняет по секциям <see cref="ComputedSections"/> те же проверки, что и
        /// ComputeSection/FillAccountRow/ParseFormula сделали бы во время расчёта, но СРАЗУ по
        /// всем строкам и без похода в 1С — чтобы одним запуском увидеть весь список проблем
        /// mapping.json, а не гоняться за ними по одной через повторные COM-подключения.
        /// </summary>
        private static List<string> ValidateMapping(MappingRoot mapping)
        {
            var problems = new List<string>();

            foreach (string sectionName in ComputedSections)
            {
                MappingSection section = mapping.Sections[sectionName];

                var ids = new HashSet<string>();
                foreach (MappingRow row in section.Rows)
                {
                    if (row.Kind != "header" && !string.IsNullOrEmpty(row.RedniBroj))
                        ids.Add(row.RedniBroj);
                }

                var seen = new HashSet<string>();
                foreach (MappingRow row in section.Rows)
                {
                    if (row.Kind == "header") continue;

                    if (string.IsNullOrEmpty(row.RedniBroj))
                    {
                        problems.Add($"[{sectionName}] строка kind='{row.Kind}' (pozicija='{row.Pozicija}') без redniBroj.");
                        continue;
                    }

                    if (!seen.Add(row.RedniBroj))
                        problems.Add($"[{sectionName}] {row.RedniBroj}: повторяющийся redniBroj.");

                    switch (row.Kind)
                    {
                        case "account":
                        {
                            AccountsSpec? accounts = row.Accounts;
                            if (accounts is null || accounts.Include is null || accounts.Include.Count == 0)
                                // Пустой include — легитимно: dio-дефолты apply_defaults.py
                                // намеренно обнуляют вторичные строки группы, сумма уходит в первую.
                                Console.Error.WriteLine(
                                    $"WARN [{sectionName}] {row.RedniBroj}: accounts.include пуст — строка считается нулевой " +
                                    $"(dio-дефолт, grupaRacuna='{row.GrupaRacuna}').");

                            if (string.IsNullOrEmpty(row.Balance))
                            {
                                problems.Add($"[{sectionName}] {row.RedniBroj}: нет поля balance.");
                            }
                            else if (row.Balance is not ("debitNet" or "creditNet" or "debitOnly" or "creditOnly"
                                         or "turnoverDebitNet" or "turnoverCreditNet"))
                            {
                                problems.Add($"[{sectionName}] {row.RedniBroj}: неизвестный тип balance='{row.Balance}'.");
                            }
                            break;
                        }
                        case "formula":
                        {
                            string? formulaText = row.Formula;
                            if (string.IsNullOrWhiteSpace(formulaText))
                            {
                                problems.Add($"[{sectionName}] {row.RedniBroj}: формула пуста.");
                                break;
                            }

                            List<FormulaTerm> terms;
                            try
                            {
                                terms = ParseFormula(formulaText, sectionName, row.RedniBroj);
                            }
                            catch (Exception ex)
                            {
                                problems.Add($"[{sectionName}] {row.RedniBroj}: {ex.Message}");
                                break;
                            }

                            foreach (FormulaTerm term in terms)
                            {
                                if (!ids.Contains(term.Reference))
                                    problems.Add(
                                        $"[{sectionName}] {row.RedniBroj}: формула ссылается на несуществующий redniBroj '{term.Reference}'.");
                            }
                            break;
                        }
                        case "manual":
                            break;
                        default:
                            problems.Add($"[{sectionName}] {row.RedniBroj}: неизвестный kind='{row.Kind}'.");
                            break;
                    }
                }
            }

            return problems;
        }

        private static int Run(object connObj, string db, string orgName, int year, MappingRoot mapping, string outJsonPath)
        {
            object? orgRef = FindOrganization(connObj, orgName);
            if (orgRef is null)
            {
                Console.Error.WriteLine($"Организация не найдена в базе '{db}': '{orgName}'.");
                return 1;
            }

            // Остатки берём на конец дня 31.12 (включительно) — как указано в задании.
            DateTime endOfYear = new DateTime(year, 12, 31, 23, 59, 59);
            DateTime endOfPrevYear = new DateTime(year - 1, 12, 31, 23, 59, 59);
            DateTime startOfYear = new DateTime(year, 1, 1, 0, 0, 0);
            DateTime startOfPrevYear = new DateTime(year - 1, 1, 1, 0, 0, 0);

            Dictionary<string, AccountAmounts> remaindersYear = LeafOnly(QueryRemainders(connObj, orgRef, endOfYear));
            Dictionary<string, AccountAmounts> remaindersPrevYear = LeafOnly(QueryRemainders(connObj, orgRef, endOfPrevYear));
            Dictionary<string, AccountAmounts> turnoversYear = LeafOnly(QueryTurnovers(connObj, orgRef, startOfYear, endOfYear));
            Dictionary<string, AccountAmounts> turnoversPrevYear = LeafOnly(QueryTurnovers(connObj, orgRef, startOfPrevYear, endOfPrevYear));

            var facts = new SectionFacts(remaindersYear, remaindersPrevYear, turnoversYear, turnoversPrevYear);

            Dictionary<string, decimal[]> bilanStanja = ComputeSection(mapping.Sections["BilanStanja"], 3, facts, "BilanStanja");
            Dictionary<string, decimal[]> bilanUspjeha = ComputeSection(mapping.Sections["BilanUspjeha"], 2, facts, "BilanUspjeha");
            Dictionary<string, decimal[]> statistickiAneks = ComputeSection(mapping.Sections["StatistickiAneks"], 2, facts, "StatistickiAneks");

            RoundInPlace(bilanStanja);
            RoundInPlace(bilanUspjeha);
            RoundInPlace(statistickiAneks);

            WriteJson(outJsonPath, orgName, year, bilanStanja, bilanUspjeha, statistickiAneks);
            PrintControlLines(bilanStanja, bilanUspjeha);

            return 0;
        }

        #region Запросы к 1С

        // ЛОВУШКА CS1977: методы ниже принимают COM-соединение как object (не dynamic) и
        // приводят его к dynamic только ЛОКАЛЬНО, внутри тела метода — так вызовы,
        // принимающие делегаты/лямбды (LINQ и т.п.), никогда не оказываются частью
        // динамически диспетчеризуемого выражения.

        private static object? FindOrganization(object connObj, string orgName)
        {
            dynamic conn = connObj;
            dynamic q = conn.NewObject("Запрос");
            q.Текст = "ВЫБРАТЬ Ссылка ИЗ Справочник.Организации ГДЕ Наименование = &Имя";
            q.УстановитьПараметр("Имя", orgName);

            dynamic result = q.Выполнить();
            dynamic sel = result.Выбрать();

            object? found = null;
            int count = 0;
            while (sel.Следующий())
            {
                count++;
                if (found is null) found = (object)sel.Ссылка;
            }

            if (count > 1)
                Console.Error.WriteLine($"ПРЕДУПРЕЖДЕНИЕ: организация '{orgName}' найдена {count} раз(а) в справочнике — используется первая найденная.");

            return found;
        }

        private static List<AccountAmounts> QueryRemainders(object connObj, object orgRef, DateTime momentEnd)
        {
            dynamic conn = connObj;
            dynamic q = conn.NewObject("Запрос");
            q.Текст = @"ВЫБРАТЬ
                    Т.Счет.Код КАК Код,
                    Т.СуммаВВалютеОтчетностиОстатокДт КАК Дт,
                    Т.СуммаВВалютеОтчетностиОстатокКт КАК Кт
                ИЗ РегистрБухгалтерии.МСФО.Остатки(&МоментКонца, , , Организация = &Орг И Сценарий.Наименование = ""Факт"") КАК Т";
            q.УстановитьПараметр("МоментКонца", momentEnd);
            q.УстановитьПараметр("Орг", orgRef);

            return ReadCodeDebitCredit(q);
        }

        private static List<AccountAmounts> QueryTurnovers(object connObj, object orgRef, DateTime start, DateTime end)
        {
            dynamic conn = connObj;
            dynamic q = conn.NewObject("Запрос");
            q.Текст = @"ВЫБРАТЬ
                    Т.Счет.Код КАК Код,
                    Т.СуммаВВалютеОтчетностиОборотДт КАК Дт,
                    Т.СуммаВВалютеОтчетностиОборотКт КАК Кт
                ИЗ РегистрБухгалтерии.МСФО.Обороты(&Нач, &Кон, , , , Организация = &Орг И Сценарий.Наименование = ""Факт"", , ) КАК Т";
            q.УстановитьПараметр("Нач", start);
            q.УстановитьПараметр("Кон", end);
            q.УстановитьПараметр("Орг", orgRef);

            return ReadCodeDebitCredit(q);
        }

        /// <summary>
        /// Общее чтение результата обоих запросов (Остатки/Обороты) — оба выбирают одноимённые
        /// колонки Код/Дт/Кт (для оборотов это СуммаОборотДт/Кт, для остатков — СуммаОстатокДт/Кт).
        /// </summary>
        private static List<AccountAmounts> ReadCodeDebitCredit(dynamic q)
        {
            dynamic result = q.Выполнить();
            dynamic sel = result.Выбрать();

            var list = new List<AccountAmounts>();
            while (sel.Следующий())
            {
                string code;
                decimal debit;
                decimal credit;
                try
                {
                    code = ((string)(sel.Код ?? "")).Trim();
                    debit = (decimal)sel.Дт;
                    credit = (decimal)sel.Кт;
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        $"Не удалось прочитать строку выборки Остатки/Обороты (неожиданный тип поля Код/Дт/Кт): {ex.Message}", ex);
                }

                if (code.Length == 0) continue; // защитный пропуск — по счёту всегда должен быть код
                list.Add(new AccountAmounts(code, debit, credit));
            }

            return list;
        }

        #endregion

        #region Листовые счета

        /// <summary>
        /// Оставляет только "листовые" коды счетов: код X выбрасывается, если в ТОЙ ЖЕ выборке
        /// есть другой, более длинный код Y, начинающийся с X (счёт-группа задваивает субсчёт).
        /// Правило применяется независимо к каждой из 4 выборок (Остатки/Обороты × Год/Год-1).
        /// </summary>
        private static Dictionary<string, AccountAmounts> LeafOnly(List<AccountAmounts> rows)
        {
            // Защитный шаг: если код в выборке почему-то повторился — суммируем (не должно
            // происходить, т.к. Организация зафиксирована условием, а Счёт — единственное
            // измерение регистра сверх Организации).
            var byCode = new Dictionary<string, AccountAmounts>();
            foreach (AccountAmounts row in rows)
            {
                if (byCode.TryGetValue(row.Code, out AccountAmounts existing))
                    byCode[row.Code] = new AccountAmounts(row.Code, existing.Debit + row.Debit, existing.Credit + row.Credit);
                else
                    byCode[row.Code] = row;
            }

            List<string> codes = byCode.Keys.ToList();
            var result = new Dictionary<string, AccountAmounts>();

            foreach (string x in codes)
            {
                bool hasLongerDescendant = false;
                foreach (string y in codes)
                {
                    if (y.Length > x.Length && y.StartsWith(x, StringComparison.Ordinal))
                    {
                        hasLongerDescendant = true;
                        break;
                    }
                }
                if (!hasLongerDescendant) result[x] = byCode[x];
            }

            return result;
        }

        #endregion

        #region Агрегация секции по mapping.json

        private static Dictionary<string, decimal[]> ComputeSection(
            MappingSection section, int columnCount, SectionFacts facts, string sectionName)
        {
            var values = new Dictionary<string, decimal[]>();
            var formulas = new List<(string RedniBroj, List<FormulaTerm> Terms)>();

            foreach (MappingRow row in section.Rows)
            {
                if (row.Kind == "header") continue;

                if (string.IsNullOrEmpty(row.RedniBroj))
                    throw new InvalidOperationException(
                        $"[{sectionName}] строка kind='{row.Kind}' (pozicija='{row.Pozicija}') без redniBroj.");

                if (values.ContainsKey(row.RedniBroj))
                    throw new InvalidOperationException($"[{sectionName}] повторяющийся redniBroj '{row.RedniBroj}'.");

                var arr = new decimal[columnCount];

                switch (row.Kind)
                {
                    case "account":
                    {
                        FillAccountRow(row, facts, columnCount, arr, sectionName);
                        break;
                    }
                    case "formula":
                    {
                        string? formulaText = row.Formula;
                        if (string.IsNullOrWhiteSpace(formulaText))
                            throw new InvalidOperationException($"[{sectionName}] формульная строка '{row.RedniBroj}' без поля formula.");
                        formulas.Add((row.RedniBroj, ParseFormula(formulaText, sectionName, row.RedniBroj)));
                        break;
                    }
                    case "manual":
                    {
                        // arr уже заполнен нулями — ручные строки в FiCalc всегда 0 (в реальной
                        // обработке 1С заполняются вручную/через ТЧ "Корректировки").
                        break;
                    }
                    default:
                        throw new InvalidOperationException($"[{sectionName}] строка '{row.RedniBroj}' имеет неизвестный kind='{row.Kind}'.");
                }

                values[row.RedniBroj] = arr;
            }

            ResolveFormulas(values, formulas, columnCount, sectionName);

            return values;
        }

        private static void FillAccountRow(MappingRow row, SectionFacts facts, int columnCount, decimal[] arr, string sectionName)
        {
            if (string.IsNullOrEmpty(row.Balance))
                throw new InvalidOperationException($"[{sectionName}] строка '{row.RedniBroj}' kind='account' без поля balance.");

            AccountsSpec? accounts = row.Accounts;
            if (accounts is null || accounts.Include is null || accounts.Include.Count == 0)
                return; // dio-дефолт: строка намеренно нулевая, сумма группы уходит в первую строку

            string balanceKind = row.Balance;
            List<string> include = accounts.Include;
            List<string> exclude = accounts.Exclude ?? EmptyStrings;

            arr[0] = AggregateAccountRow(balanceKind, include, exclude, facts.Select(balanceKind, previousYear: false));
            if (columnCount >= 2)
                arr[1] = AggregateAccountRow(balanceKind, include, exclude, facts.Select(balanceKind, previousYear: true));
            // arr[2] (только у BilanStanja) остаётся 0 — см. комментарий в шапке файла.
        }

        /// <summary>
        /// Агрегирует один факт-набор (остатки ЛИБО обороты одного года) по правилам одной строки:
        ///   debitNet         = Σ(Дт) − Σ(Кт)
        ///   creditNet        = Σ(Кт) − Σ(Дт)
        ///   debitOnly        = Σ max(Дт−Кт, 0) по каждому счёту
        ///   creditOnly       = Σ max(Кт−Дт, 0) по каждому счёту
        ///   turnoverDebitNet = Σ(ОборотДт) − Σ(ОборотКт)   (те же Дт/Кт, но из выборки Обороты)
        ///   turnoverCreditNet= Σ(ОборотКт) − Σ(ОборотДт)
        /// </summary>
        private static decimal AggregateAccountRow(
            string balanceKind, List<string> include, List<string> exclude, Dictionary<string, AccountAmounts> factsForColumn)
        {
            decimal debitNet = 0m;
            decimal creditNet = 0m;
            decimal debitOnly = 0m;
            decimal creditOnly = 0m;

            foreach (KeyValuePair<string, AccountAmounts> kv in factsForColumn)
            {
                string code = kv.Key;
                if (!StartsWithAny(code, include)) continue;
                if (StartsWithAny(code, exclude)) continue;

                decimal net = kv.Value.Debit - kv.Value.Credit; // Дт-Кт (для оборотов: ОборотДт-ОборотКт)
                debitNet += net;
                creditNet -= net;
                if (net > 0m) debitOnly += net;
                else if (net < 0m) creditOnly += -net;
            }

            return balanceKind switch
            {
                "debitNet" => debitNet,
                "creditNet" => creditNet,
                "debitOnly" => debitOnly,
                "creditOnly" => creditOnly,
                "turnoverDebitNet" => debitNet,
                "turnoverCreditNet" => creditNet,
                _ => throw new InvalidOperationException($"Неизвестный тип balance: '{balanceKind}'.")
            };
        }

        private static bool StartsWithAny(string code, List<string> prefixes)
        {
            foreach (string p in prefixes)
            {
                if (code.StartsWith(p, StringComparison.Ordinal)) return true;
            }
            return false;
        }

        #endregion

        #region Формулы

        /// <summary>
        /// Разбирает формулу вида "003+008+016" или "106+107+108+109-110" (без пробелов, ссылки
        /// могут иметь буквенный суффикс, например "210a") в список слагаемых со знаками.
        /// </summary>
        private static List<FormulaTerm> ParseFormula(string formula, string sectionName, string redniBroj)
        {
            string compact = formula.Replace(" ", "").Replace("\t", "");
            if (compact.Length == 0)
                throw new InvalidOperationException($"[{sectionName}] формула строки '{redniBroj}' пуста.");

            MatchCollection matches = FormulaTokenRegex.Matches(compact);

            int consumedLength = 0;
            foreach (Match m in matches) consumedLength += m.Length;

            if (matches.Count == 0 || consumedLength != compact.Length)
                throw new InvalidOperationException(
                    $"[{sectionName}] не удалось разобрать формулу '{formula}' строки '{redniBroj}' (нераспознанные символы).");

            var terms = new List<FormulaTerm>(matches.Count);
            foreach (Match m in matches)
            {
                string t = m.Value;
                bool negative = t[0] == '-';
                string reference = (t[0] == '+' || t[0] == '-') ? t.Substring(1) : t;
                if (reference.Length == 0)
                    throw new InvalidOperationException($"[{sectionName}] пустая ссылка в формуле '{formula}' строки '{redniBroj}'.");
                terms.Add(new FormulaTerm(negative, reference));
            }
            return terms;
        }

        /// <summary>
        /// Вычисляет формульные строки итеративно до неподвижной точки: формулы могут ссылаться
        /// на другие формулы (например BilanStanja 025=026+031+..., где 026 сама формула), поэтому
        /// простой топологической сортировки нет смысла городить — несколько проходов дают тот же
        /// результат и надёжно ловят циклические ссылки как ошибку.
        /// </summary>
        private static void ResolveFormulas(
            Dictionary<string, decimal[]> values,
            List<(string RedniBroj, List<FormulaTerm> Terms)> formulas,
            int columnCount,
            string sectionName)
        {
            if (formulas.Count == 0) return;

            int maxPasses = formulas.Count + 10;
            for (int pass = 0; pass < maxPasses; pass++)
            {
                bool changed = false;

                foreach (var (redniBroj, terms) in formulas)
                {
                    decimal[] arr = values[redniBroj];

                    for (int col = 0; col < columnCount; col++)
                    {
                        decimal newValue = 0m;
                        foreach (FormulaTerm term in terms)
                        {
                            if (!values.TryGetValue(term.Reference, out var refArr))
                                throw new InvalidOperationException(
                                    $"[{sectionName}] формула строки '{redniBroj}' ссылается на несуществующий redniBroj '{term.Reference}'.");
                            decimal v = refArr[col];
                            newValue += term.Negative ? -v : v;
                        }

                        if (newValue != arr[col])
                        {
                            arr[col] = newValue;
                            changed = true;
                        }
                    }
                }

                if (!changed) return;

                if (pass == maxPasses - 1)
                    throw new InvalidOperationException(
                        $"[{sectionName}] формулы не сошлись за {maxPasses} проходов — возможна циклическая ссылка.");
            }
        }

        #endregion

        #region Округление, вывод

        private static void RoundInPlace(Dictionary<string, decimal[]> values)
        {
            foreach (decimal[] arr in values.Values)
            {
                for (int i = 0; i < arr.Length; i++)
                    arr[i] = Math.Round(arr[i], 2, MidpointRounding.AwayFromZero);
            }
        }

        private static void WriteJson(
            string outJsonPath, string orgName, int year,
            Dictionary<string, decimal[]> bilanStanja,
            Dictionary<string, decimal[]> bilanUspjeha,
            Dictionary<string, decimal[]> statistickiAneks)
        {
            var payload = new FiCalcOutput
            {
                Org = orgName,
                Year = year,
                Sections = new Dictionary<string, Dictionary<string, decimal[]>>
                {
                    ["BilanStanja"] = bilanStanja,
                    ["BilanUspjeha"] = bilanUspjeha,
                    ["StatistickiAneks"] = statistickiAneks,
                },
                Meta = new FiCalcMeta
                {
                    Columns = new Dictionary<string, string>
                    {
                        ["BilanStanja"] = "Iznos1=Y,Iznos2=Y-1,Iznos3=0",
                        ["BilanUspjeha"] = "Iznos1=Y,Iznos2=Y-1",
                        ["StatistickiAneks"] = "Iznos1=Y,Iznos2=Y-1",
                    }
                }
            };

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            };

            string json = JsonSerializer.Serialize(payload, options);

            string? dir = Path.GetDirectoryName(outJsonPath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            File.WriteAllText(outJsonPath, json, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            Console.WriteLine($"JSON записан: {outJsonPath}");
        }

        private static void PrintControlLines(Dictionary<string, decimal[]> bilanStanja, Dictionary<string, decimal[]> bilanUspjeha)
        {
            Console.WriteLine();
            Console.WriteLine("=== Kontrolne linije (kolona Iznos1 = tekuca godina) ===");

            decimal aktiva = GetControlValue(bilanStanja, "046", "046 (ukupna aktiva)");
            decimal pasiva = GetControlValue(bilanStanja, "144", "144 (ukupna pasiva)");
            decimal rezultat = GetControlValue(bilanUspjeha, "244", "244 (rezultat prije oporezivanja)");
            decimal razlika = aktiva - pasiva;

            Console.WriteLine($"046 Ukupna aktiva:                 {aktiva.ToString("N2", CultureInfo.InvariantCulture)}");
            Console.WriteLine($"144 Ukupna pasiva:                  {pasiva.ToString("N2", CultureInfo.InvariantCulture)}");
            Console.WriteLine($"Razlika (046-144), treba biti 0:    {razlika.ToString("N2", CultureInfo.InvariantCulture)}");
            Console.WriteLine($"244 Rezultat prije oporezivanja:    {rezultat.ToString("N2", CultureInfo.InvariantCulture)}");
        }

        private static decimal GetControlValue(Dictionary<string, decimal[]> section, string redniBroj, string label)
        {
            if (section.TryGetValue(redniBroj, out var arr) && arr.Length > 0) return arr[0];
            Console.Error.WriteLine($"ПРЕДУПРЕЖДЕНИЕ: контрольная строка {label} не найдена в mapping.json.");
            return 0m;
        }

        #endregion
    }
}
