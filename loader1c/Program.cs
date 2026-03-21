using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

// ──────────────────────────────────────────────────────────────
// Entry point
// ──────────────────────────────────────────────────────────────
public class Program
{
    public static void Main(string[] args)
    {
        // Register windows-1251 encoding
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        var config = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false)
            .Build();

        var loaderConfig = config.GetSection("Loader").Get<LoaderConfig>()
            ?? throw new Exception("Missing 'Loader' section in appsettings.json");

        // CLI commands
        var command = args.Length > 0 ? args[0].ToLower() : "run";

        if (command == "diag")
        {
            // Diagnostic: discover property names for bank documents
            using var loggerFactory = LoggerFactory.Create(b => b.AddConsole());
            var logger = loggerFactory.CreateLogger<Program>();

            string db = args.Length > 1 ? args[1] : "base_rs";
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine($"Connecting to {db}...");

            var connectorType = Type.GetTypeFromProgID("V83.COMConnector")
                ?? throw new Exception("V83.COMConnector not registered");
            dynamic connector = Activator.CreateInstance(connectorType)!;
            var connStr = $"Srvr=\"{loaderConfig.Server}\";Ref=\"{db}\";" +
                          $"Usr=\"{loaderConfig.User}\";Pwd=\"{loaderConfig.Password}\"";
            dynamic conn = connector.Connect(connStr);
            Console.WriteLine("Connected!");

            // Create a test debit document to inspect its metadata
            Console.WriteLine("\n=== СписаниеСРасчетногоСчета ===");
            dynamic debitDoc = conn.Документы.СписаниеСРасчетногоСчета.СоздатьДокумент();
            dynamic debitMeta = debitDoc.Метаданные();

            Console.WriteLine("Реквизиты:");
            dynamic debitAttrs = debitMeta.Реквизиты;
            for (int i = 0; i < (int)debitAttrs.Количество(); i++)
            {
                dynamic attr = debitAttrs.Получить(i);
                Console.WriteLine($"  {attr.Имя}  ({attr.Тип})");
            }

            Console.WriteLine("\nТабличные части:");
            dynamic debitTables = debitMeta.ТабличныеЧасти;
            for (int i = 0; i < (int)debitTables.Количество(); i++)
            {
                dynamic table = debitTables.Получить(i);
                Console.WriteLine($"  {table.Имя}:");
                dynamic tableCols = table.Реквизиты;
                for (int j = 0; j < (int)tableCols.Количество(); j++)
                {
                    dynamic col = tableCols.Получить(j);
                    Console.WriteLine($"    {col.Имя}  ({col.Тип})");
                }
            }

            // List operation type enums
            Console.WriteLine("\n=== Перечисления (операции) ===");
            dynamic allEnums = conn.Метаданные.Перечисления;
            for (int i = 0; i < (int)allEnums.Количество(); i++)
            {
                dynamic en = allEnums.Получить(i);
                string name = en.Имя;
                if (name.Contains("Операци") || name.Contains("Списани") ||
                    name.Contains("Поступлени") || name.Contains("Банк") ||
                    name.Contains("Налог") || name.Contains("Обязательств") ||
                    name.Contains("Бюджет"))
                {
                    Console.WriteLine($"  {name}:");
                    dynamic vals = en.ЗначенияПеречисления;
                    for (int j = 0; j < (int)vals.Количество(); j++)
                    {
                        Console.WriteLine($"    {vals.Получить(j).Имя}");
                    }
                }
            }

            // List relevant catalogs
            Console.WriteLine("\n=== Справочники (движ.ден.средств, расходы) ===");
            dynamic allCats = conn.Метаданные.Справочники;
            for (int i = 0; i < (int)allCats.Количество(); i++)
            {
                dynamic cat = allCats.Получить(i);
                string catName = cat.Имя;
                if (catName.Contains("Стать") || catName.Contains("Движени") ||
                    catName.Contains("Расход") || catName.Contains("Налог") ||
                    catName.Contains("Подразделен"))
                {
                    Console.WriteLine($"  {catName}");
                }
            }

            // Dump specific catalog values
            // Check the actual type of ВидНалоговогоОбязательства attribute
            Console.WriteLine("\n=== Type of ВидНалоговогоОбязательства ===");
            try
            {
                dynamic vidNalOblAttr = debitMeta.Реквизиты.Найти("ВидНалоговогоОбязательства");
                if (vidNalOblAttr != null)
                {
                    dynamic typeDesc = vidNalOblAttr.Тип;
                    Console.WriteLine($"  TypeDescription: {typeDesc}");
                    dynamic typeArr = typeDesc.Типы();
                    int cnt = (int)typeArr.Количество();
                    Console.WriteLine($"  Types count: {cnt}");
                    for (int i = 0; i < cnt; i++)
                    {
                        dynamic t = typeArr.Получить(i);
                        Console.WriteLine($"  Type[{i}]: {t}");
                        // Try to get type name via metadata
                        try
                        {
                            dynamic md = conn.Метаданные.НайтиПоТипу(t);
                            Console.WriteLine($"    Metadata: {md.ПолноеИмя()}");
                        }
                        catch (Exception ex2) { Console.WriteLine($"    Metadata: {ex2.Message}"); }
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine($"  Error: {ex.Message}"); }

            Console.WriteLine("\n=== Catalog values: ВидыОбязательствНалогоплательщика ===");
            try
            {
                dynamic oblSel2 = conn.Справочники.ВидыОбязательствНалогоплательщика.Выбрать();
                while (oblSel2.Следующий())
                    Console.WriteLine($"  '{oblSel2.Наименование}' (deleted={oblSel2.ПометкаУдаления})");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  (catalog error: {ex.Message})");
                // Try as enum
                try
                {
                    dynamic allEnums3 = conn.Метаданные.Перечисления;
                    for (int i = 0; i < (int)allEnums3.Количество(); i++)
                    {
                        dynamic en = allEnums3.Получить(i);
                        if (en.Имя.ToString().Contains("Обязательств"))
                        {
                            Console.WriteLine($"  Enum: {en.Имя}");
                            dynamic vals = en.ЗначенияПеречисления;
                            for (int j = 0; j < (int)vals.Количество(); j++)
                                Console.WriteLine($"    {vals.Получить(j).Имя}");
                        }
                    }
                }
                catch { }
            }

            Console.WriteLine("\n=== Catalog values: ВидыНалоговыхОбязательств ===");
            try
            {
                dynamic oblSel = conn.Справочники.ВидыНалоговыхОбязательств.Выбрать();
                while (oblSel.Следующий())
                    Console.WriteLine($"  '{oblSel.Наименование}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  (catalog not found, trying enum: {ex.Message})");
                try
                {
                    dynamic allEnums2 = conn.Метаданные.Перечисления;
                    for (int i = 0; i < (int)allEnums2.Количество(); i++)
                    {
                        dynamic en = allEnums2.Получить(i);
                        if (en.Имя.ToString().Contains("Обязательств"))
                        {
                            Console.WriteLine($"  Enum: {en.Имя}");
                            dynamic vals = en.ЗначенияПеречисления;
                            for (int j = 0; j < (int)vals.Количество(); j++)
                                Console.WriteLine($"    {vals.Получить(j).Имя}");
                        }
                    }
                }
                catch { }
            }

            Console.WriteLine("\n=== Catalog values: ПодразделенияОрганизаций ===");
            try
            {
                dynamic deptSel = conn.Справочники.ПодразделенияОрганизаций.Выбрать();
                while (deptSel.Следующий())
                    Console.WriteLine($"  '{deptSel.Наименование}' (deleted={deptSel.ПометкаУдаления})");
            }
            catch (Exception ex) { Console.WriteLine($"  Error: {ex.Message}"); }

            Console.WriteLine("\n=== Catalog values: ВидыНалоговИПлатежейВБюджет ===");
            try
            {
                dynamic taxSel = conn.Справочники.ВидыНалоговИПлатежейВБюджет.Выбрать();
                while (taxSel.Следующий())
                    Console.WriteLine($"  '{taxSel.Наименование}' (deleted={taxSel.ПометкаУдаления})");
            }
            catch (Exception ex) { Console.WriteLine($"  Error: {ex.Message}"); }

            Console.WriteLine("\n=== ПланыСчетов.Хозрасчетный ===");
            string[] testCodes = { "51", "4520", "45.20", "45", "68", "68.10", "69", "69.20", "69.20.1", "91.02" };
            foreach (var code in testCodes)
            {
                try
                {
                    dynamic acct = conn.ПланыСчетов.Хозрасчетный.НайтиПоКоду(code);
                    bool isEmpty = acct.Пустая();
                    string name = isEmpty ? "(empty)" : (string)acct.Наименование;
                    Console.WriteLine($"  {code} → {name} (empty={isEmpty})");
                }
                catch (Exception ex) { Console.WriteLine($"  {code} → Error: {ex.Message}"); }
            }

            // Dump ВедомостьНаВыплатуЗарплаты metadata
            Console.WriteLine("\n=== ВедомостьНаВыплатуЗарплаты metadata ===");
            try
            {
                dynamic vedDoc = conn.Документы.ВедомостьНаВыплатуЗарплаты.СоздатьДокумент();
                dynamic vedMeta = vedDoc.Метаданные();
                dynamic vedAttrs = vedMeta.Реквизиты;
                for (int i = 0; i < (int)vedAttrs.Количество(); i++)
                {
                    dynamic attr = vedAttrs.Получить(i);
                    Console.WriteLine($"  {attr.Имя}  ({attr.Тип})");
                    try
                    {
                        dynamic typeArr = attr.Тип.Типы();
                        for (int j = 0; j < (int)typeArr.Количество(); j++)
                        {
                            dynamic t = typeArr.Получить(j);
                            try
                            {
                                dynamic md = conn.Метаданные.НайтиПоТипу(t);
                                string fullName = (string)md.ПолноеИмя();
                                Console.WriteLine($"    -> {fullName}");
                                if (fullName.StartsWith("Перечисление."))
                                {
                                    string enumName = fullName.Replace("Перечисление.", "");
                                    dynamic enumMeta = conn.Метаданные.Перечисления.Найти(enumName);
                                    if (enumMeta != null)
                                    {
                                        dynamic vals = enumMeta.ЗначенияПеречисления;
                                        for (int k = 0; k < (int)vals.Количество(); k++)
                                            Console.WriteLine($"      {vals.Получить(k).Имя}");
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Console.WriteLine($"  Error: {ex.Message}"); }

            Marshal.ReleaseComObject(conn);
            Marshal.ReleaseComObject(connector);
            return;
        }

        if (command == "inspect" && args.Length > 2)
        {
            // Inspect a specific document by number: inspect <db> <docNumber>
            using var loggerFactory = LoggerFactory.Create(b => b.AddConsole());
            var logger = loggerFactory.CreateLogger<Program>();
            string db = args[1];
            string docNum = args[2];
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine($"Connecting to {db}...");

            var connectorType = Type.GetTypeFromProgID("V83.COMConnector")
                ?? throw new Exception("V83.COMConnector not registered");
            dynamic connector = Activator.CreateInstance(connectorType)!;
            var connStr = $"Srvr=\"{loaderConfig.Server}\";Ref=\"{db}\";" +
                          $"Usr=\"{loaderConfig.User}\";Pwd=\"{loaderConfig.Password}\"";
            dynamic conn = connector.Connect(connStr);
            Console.WriteLine("Connected!");

            // Search in both document types — show ALL matches
            bool found = false;
            foreach (var docType in new[] { "СписаниеСРасчетногоСчета", "ПоступлениеНаРасчетныйСчет" })
            {
                try
                {
                    dynamic mgr = conn.Документы.GetType().InvokeMember(
                        docType, System.Reflection.BindingFlags.GetProperty,
                        null, (object)conn.Документы, null);
                    dynamic sel = mgr.Выбрать();
                    while (sel.Следующий())
                    {
                        string num = (string)(sel.Номер ?? "");
                        if (!num.Contains(docNum.Trim(), StringComparison.OrdinalIgnoreCase)) continue;
                        found = true;

                        // Get object for full attribute access
                        dynamic docObj = sel.Ссылка.ПолучитьОбъект();
                        Console.WriteLine($"\n=== {docType} #{num} от {sel.Дата} ===");

                        // Dump ALL attributes via metadata
                        dynamic meta = docObj.Метаданные();
                        dynamic attrs = meta.Реквизиты;
                        for (int i = 0; i < (int)attrs.Количество(); i++)
                        {
                            dynamic attr = attrs.Получить(i);
                            string attrName = (string)attr.Имя;
                            try
                            {
                                object val = docObj.GetType().InvokeMember(
                                    attrName, System.Reflection.BindingFlags.GetProperty,
                                    null, (object)docObj, null);
                                string valStr = val?.ToString() ?? "(null)";
                                // For COM objects try to get Наименование or НомерСчета
                                if (valStr == "System.__ComObject")
                                {
                                    try { valStr = ((dynamic)val).Наименование ?? valStr; } catch {
                                        try { valStr = ((dynamic)val).НомерСчета ?? valStr; } catch {
                                            try { valStr = ((dynamic)val).Код ?? valStr; } catch {
                                                try { valStr = $"#{((dynamic)val).Номер} от {((dynamic)val).Дата}"; } catch { }
                                            }
                                        }
                                    }
                                    // For ПлатежнаяВедомость — dump its attributes too
                                    if (attrName == "ПлатежнаяВедомость" && valStr != "(null)")
                                    {
                                        Console.WriteLine($"  {attrName} = {valStr}");
                                        try
                                        {
                                            dynamic vedObj = ((dynamic)val).ПолучитьОбъект();
                                            dynamic vedMeta = vedObj.Метаданные();
                                            Console.WriteLine($"    >>> Тип: {vedMeta.ПолноеИмя()}");
                                            dynamic vedAttrs = vedMeta.Реквизиты;
                                            for (int va = 0; va < (int)vedAttrs.Количество(); va++)
                                            {
                                                dynamic vattr = vedAttrs.Получить(va);
                                                string vName = (string)vattr.Имя;
                                                try
                                                {
                                                    object vv = vedObj.GetType().InvokeMember(
                                                        vName, System.Reflection.BindingFlags.GetProperty,
                                                        null, (object)vedObj, null);
                                                    string vvStr = vv?.ToString() ?? "(null)";
                                                    if (vvStr == "System.__ComObject")
                                                        try { vvStr = ((dynamic)vv).Наименование ?? vvStr; } catch {
                                                            try { vvStr = $"#{((dynamic)vv).Номер} от {((dynamic)vv).Дата}"; } catch { }
                                                        }
                                                    if (vvStr.Length > 100) vvStr = vvStr[..100] + "...";
                                                    Console.WriteLine($"    {vName} = {vvStr}");
                                                }
                                                catch { }
                                            }
                                            // Tabular sections of the vedомость
                                            dynamic vedTables = vedMeta.ТабличныеЧасти;
                                            for (int vt = 0; vt < (int)vedTables.Количество(); vt++)
                                            {
                                                dynamic vtable = vedTables.Получить(vt);
                                                string vtName = (string)vtable.Имя;
                                                try
                                                {
                                                    dynamic vtSection = vedObj.GetType().InvokeMember(
                                                        vtName, System.Reflection.BindingFlags.GetProperty,
                                                        null, (object)vedObj, null);
                                                    int vtRows = (int)vtSection.Количество();
                                                    Console.WriteLine($"    --- {vtName}: {vtRows} rows ---");
                                                    dynamic vtCols = vtable.Реквизиты;
                                                    int vtColCount = (int)vtCols.Количество();
                                                    for (int vr = 0; vr < Math.Min(vtRows, 10); vr++)
                                                    {
                                                        dynamic vrow = vtSection.Получить(vr);
                                                        Console.Write($"      row[{vr}]: ");
                                                        for (int vc = 0; vc < vtColCount; vc++)
                                                        {
                                                            string vcName = (string)vtCols.Получить(vc).Имя;
                                                            try
                                                            {
                                                                object vcv = vrow.GetType().InvokeMember(
                                                                    vcName, System.Reflection.BindingFlags.GetProperty,
                                                                    null, (object)vrow, null);
                                                                string vcvStr = vcv?.ToString() ?? "";
                                                                if (vcvStr == "System.__ComObject")
                                                                    try { vcvStr = ((dynamic)vcv).Наименование ?? vcvStr; } catch { }
                                                                Console.Write($"{vcName}={vcvStr}; ");
                                                            }
                                                            catch { }
                                                        }
                                                        Console.WriteLine();
                                                    }
                                                }
                                                catch { }
                                            }
                                        }
                                        catch (Exception vex) { Console.WriteLine($"    >>> Error: {vex.Message}"); }
                                        continue; // already printed
                                    }
                                }
                                if (valStr.Length > 120) valStr = valStr[..120] + "...";
                                Console.WriteLine($"  {attrName} = {valStr}");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"  {attrName} = ERROR: {ex.Message}");
                            }
                        }

                        // Tabular sections
                        dynamic tables = meta.ТабличныеЧасти;
                        for (int ti = 0; ti < (int)tables.Количество(); ti++)
                        {
                            dynamic table = tables.Получить(ti);
                            string tableName = (string)table.Имя;
                            try
                            {
                                dynamic tabSection = docObj.GetType().InvokeMember(
                                    tableName, System.Reflection.BindingFlags.GetProperty,
                                    null, (object)docObj, null);
                                int rowCount = (int)tabSection.Количество();
                                Console.WriteLine($"\n  --- {tableName}: {rowCount} rows ---");
                                if (rowCount > 0)
                                {
                                    dynamic tableCols = table.Реквизиты;
                                    int colCount = (int)tableCols.Количество();
                                    for (int r = 0; r < Math.Min(rowCount, 5); r++)
                                    {
                                        dynamic row = tabSection.Получить(r);
                                        Console.Write($"    row[{r}]: ");
                                        for (int c = 0; c < colCount; c++)
                                        {
                                            string colName = (string)tableCols.Получить(c).Имя;
                                            try
                                            {
                                                object cv = row.GetType().InvokeMember(
                                                    colName, System.Reflection.BindingFlags.GetProperty,
                                                    null, (object)row, null);
                                                string cvStr = cv?.ToString() ?? "";
                                                if (cvStr == "System.__ComObject")
                                                    try { cvStr = ((dynamic)cv).Наименование ?? cvStr; } catch { }
                                                Console.Write($"{colName}={cvStr}; ");
                                            }
                                            catch { Console.Write($"{colName}=?; "); }
                                        }
                                        Console.WriteLine();
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch { }
            }
            if (!found) Console.WriteLine($"Document {docNum} not found in {db}");
            Marshal.ReleaseComObject(conn);
            Marshal.ReleaseComObject(connector);
            return;
        }

        if (command == "fix-bankacct" && args.Length > 2)
        {
            // Fix bank account attributes: fix-bankacct <db> <account>
            string db = args[1];
            string acctNum = args[2];
            Console.OutputEncoding = Encoding.UTF8;
            var connectorType = Type.GetTypeFromProgID("V83.COMConnector")
                ?? throw new Exception("V83.COMConnector not registered");
            dynamic connector = Activator.CreateInstance(connectorType)!;
            var connStr = $"Srvr=\"{loaderConfig.Server}\";Ref=\"{db}\";" +
                          $"Usr=\"{loaderConfig.User}\";Pwd=\"{loaderConfig.Password}\"";
            dynamic conn = connector.Connect(connStr);

            string normalized = acctNum.Replace("-", "").Replace(" ", "");
            dynamic sel = conn.Справочники.БанковскиеСчета.Выбрать();
            bool found = false;
            while (sel.Следующий())
            {
                string no = ((string)(sel.НомерСчета ?? "")).Replace("-", "").Replace(" ", "");
                if (no != normalized) continue;
                found = true;
                dynamic obj = sel.Ссылка.ПолучитьОбъект();

                // ВидСчета
                try
                {
                    string cur = obj.ВидСчета?.ToString() ?? "";
                    if (string.IsNullOrWhiteSpace(cur) || cur == "System.__ComObject")
                    {
                        // Try known enum names
                        string[] enumNames = { "ВидыСчетов", "ВидыСчетовБанковскихСчетов",
                            "ВидыСчетовБанка", "ВидСчета" };
                        bool set = false;
                        foreach (var en in enumNames)
                        {
                            try
                            {
                                dynamic enumMgr = conn.Перечисления.GetType().InvokeMember(en,
                                    System.Reflection.BindingFlags.GetProperty, null, (object)conn.Перечисления, null);
                                var val = enumMgr.GetType().InvokeMember("Расчетный",
                                    System.Reflection.BindingFlags.GetProperty, null, (object)enumMgr, null);
                                obj.ВидСчета = val;
                                Console.WriteLine($"  Set ВидСчета = Расчетный (via {en})");
                                set = true;
                                break;
                            }
                            catch { }
                        }
                        if (!set)
                        {
                            // Try reading the value from an existing good account
                            dynamic goodSel = conn.Справочники.БанковскиеСчета.Выбрать();
                            while (goodSel.Следующий())
                            {
                                try
                                {
                                    if ((bool)goodSel.ПометкаУдаления) continue;
                                    string gno = ((string)(goodSel.НомерСчета ?? "")).Trim();
                                    if (gno == normalized) continue;
                                    dynamic goodObj = goodSel.Ссылка.ПолучитьОбъект();
                                    dynamic goodVid = goodObj.ВидСчета;
                                    if (goodVid != null && goodVid.ToString() != "" && goodVid.ToString() != "System.__ComObject")
                                    {
                                        obj.ВидСчета = goodVid;
                                        Console.WriteLine($"  Set ВидСчета from existing: {goodVid}");
                                        set = true;
                                        break;
                                    }
                                }
                                catch { }
                            }
                        }
                        if (!set) Console.WriteLine("  Could not set ВидСчета");
                    }
                    else Console.WriteLine($"  ВидСчета already set: {cur}");
                }
                catch (Exception ex) { Console.WriteLine($"  ВидСчета error: {ex.Message}"); }

                // Банк
                try
                {
                    string bankName = "";
                    try { bankName = (string)(obj.Банк.Наименование ?? ""); } catch { }
                    if (string.IsNullOrWhiteSpace(bankName))
                    {
                        string bankCode = normalized.Length >= 3 ? normalized[..3] : "";
                        dynamic bankSel = conn.Справочники.Банки.Выбрать();
                        while (bankSel.Следующий())
                        {
                            string code = ((string)(bankSel.Код ?? "")).Trim();
                            if (code == bankCode)
                            {
                                obj.Банк = bankSel.Ссылка;
                                Console.WriteLine($"  Set Банк = {bankSel.Наименование} ({bankCode})");
                                break;
                            }
                        }
                    }
                    else Console.WriteLine($"  Банк already set: {bankName}");
                }
                catch (Exception ex) { Console.WriteLine($"  Банк error: {ex.Message}"); }

                // Валюта
                try
                {
                    string curName = "";
                    try { curName = (string)(obj.ВалютаДенежныхСредств.Наименование ?? ""); } catch { }
                    if (string.IsNullOrWhiteSpace(curName))
                    {
                        dynamic eur = conn.Справочники.Валюты.НайтиПоНаименованию("EUR", true);
                        if (!eur.Пустая())
                        {
                            obj.ВалютаДенежныхСредств = eur;
                            Console.WriteLine("  Set Валюта = EUR");
                        }
                    }
                    else Console.WriteLine($"  Валюта already set: {curName}");
                }
                catch (Exception ex) { Console.WriteLine($"  Валюта error: {ex.Message}"); }

                obj.Записать();
                Console.WriteLine("  Saved OK");
                break;
            }
            if (!found) Console.WriteLine($"Bank account {acctNum} not found");
            Marshal.ReleaseComObject(conn);
            Marshal.ReleaseComObject(connector);
            return;
        }

        if (command == "bankacct-meta" && args.Length > 1)
        {
            // Dump bank account metadata + sample values
            string db = args[1];
            Console.OutputEncoding = Encoding.UTF8;
            var connectorType = Type.GetTypeFromProgID("V83.COMConnector")
                ?? throw new Exception("V83.COMConnector not registered");
            dynamic connector = Activator.CreateInstance(connectorType)!;
            var connStr = $"Srvr=\"{loaderConfig.Server}\";Ref=\"{db}\";" +
                          $"Usr=\"{loaderConfig.User}\";Pwd=\"{loaderConfig.Password}\"";
            dynamic conn = connector.Connect(connStr);

            // Show first 5 existing bank accounts with ВидСчета
            Console.WriteLine("=== Sample bank accounts ===");
            dynamic sel = conn.Справочники.БанковскиеСчета.Выбрать();
            int cnt = 0;
            while (sel.Следующий() && cnt < 5)
            {
                if ((bool)sel.ПометкаУдаления) continue;
                string no = (string)(sel.НомерСчета ?? "");
                string vidSchetaStr = "";
                try { vidSchetaStr = sel.ВидСчета.ToString(); } catch { vidSchetaStr = "?"; }
                Console.WriteLine($"  {no}  ВидСчета={vidSchetaStr}");
                cnt++;
            }
            Marshal.ReleaseComObject(conn);
            Marshal.ReleaseComObject(connector);
            return;
        }

        if (command == "vat" && args.Length > 1)
        {
            string db = args[1];
            var connectorType = Type.GetTypeFromProgID("V83.COMConnector")
                ?? throw new Exception("V83.COMConnector not registered");
            dynamic connector = Activator.CreateInstance(connectorType)!;
            var connStr = $"Srvr=\"{loaderConfig.Server}\";Ref=\"{db}\";" +
                          $"Usr=\"{loaderConfig.User}\";Pwd=\"{loaderConfig.Password}\"";
            dynamic conn = connector.Connect(connStr);
            Console.OutputEncoding = Encoding.UTF8;
            // List НДС rates
            Console.WriteLine("=== СтавкиНДС ===");
            try
            {
                foreach (var name in new[] { "НДС21", "НДС20", "НДС18", "НДС10", "БезНДС",
                    "НДС_21", "НДС_20", "НДС0", "НДС5", "НДС7", "НДС15",
                    "Ставка21", "Ставка_21", "_21", "Процент21",
                    "НДС20_21", "НДС21Процент", "Ставка21Процент" })
                {
                    try
                    {
                        dynamic enumMgr = conn.Перечисления.СтавкиНДС;
                        var val = enumMgr.GetType().InvokeMember(name,
                            System.Reflection.BindingFlags.GetProperty, null, (object)enumMgr, null);
                        Console.WriteLine($"  FOUND: {name} = {val}");
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Console.WriteLine($"Enum error: {ex.Message}"); }
            // ВидыОперацийПоступлениеДенежныхСредств — all values
            Console.WriteLine("=== ВидыОперацийПоступлениеДенежныхСредств ===");
            foreach (var name in new[] { "ПереводСДругогоСчета", "ПереводНаДругойСчет",
                "ПоступлениеСДругогоСчета", "ВозвратИзДругогоБанка", "ПереводМеждуСчетами",
                "ОплатаПокупателя", "ПрочееПоступление", "ВозвратНалога",
                "ПоступлениеОтДругойОрганизации", "ПереводОтДругойОрганизации" })
            {
                try
                {
                    dynamic enumMgr = conn.Перечисления.ВидыОперацийПоступлениеДенежныхСредств;
                    var val = enumMgr.GetType().InvokeMember(name,
                        System.Reflection.BindingFlags.GetProperty, null, (object)enumMgr, null);
                    Console.WriteLine($"  FOUND: {name}");
                }
                catch { }
            }
            // СпособыПогашенияЗадолженности
            Console.WriteLine("=== СпособыПогашенияЗадолженности ===");
            try
            {
                foreach (var name in new[] { "Автоматически", "ПоДокументу", "ПоЗадолженности",
                    "ПоОстаткамВзаиморасчетов", "Auto" })
                {
                    try
                    {
                        dynamic enumMgr = conn.Перечисления.СпособыПогашенияЗадолженности;
                        var val = enumMgr.GetType().InvokeMember(name,
                            System.Reflection.BindingFlags.GetProperty, null, (object)enumMgr, null);
                        Console.WriteLine($"  FOUND: {name}");
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Console.WriteLine($"  Enum error: {ex.Message}"); }
            // Also check справочник if exists
            try
            {
                Console.WriteLine("=== Справочник СтавкиНДС ===");
                dynamic sel = conn.Справочники.СтавкиНДС.Выбрать();
                int cnt = 0;
                while (sel.Следующий() && cnt < 10)
                {
                    Console.WriteLine($"  {sel.Наименование} (code={sel.Код})");
                    cnt++;
                }
            }
            catch { Console.WriteLine("  (not a catalog)"); }
            Marshal.ReleaseComObject(conn);
            Marshal.ReleaseComObject(connector);
            return;
        }

        if (command == "banks" && args.Length > 1)
        {
            // Dump all banks
            string db = args[1];
            Console.OutputEncoding = Encoding.UTF8;
            var connectorType = Type.GetTypeFromProgID("V83.COMConnector")
                ?? throw new Exception("V83.COMConnector not registered");
            dynamic connector = Activator.CreateInstance(connectorType)!;
            var connStr = $"Srvr=\"{loaderConfig.Server}\";Ref=\"{db}\";" +
                          $"Usr=\"{loaderConfig.User}\";Pwd=\"{loaderConfig.Password}\"";
            dynamic conn = connector.Connect(connStr);
            dynamic sel = conn.Справочники.Банки.Выбрать();
            int count = 0;
            while (sel.Следующий())
            {
                if ((bool)sel.ПометкаУдаления) continue;
                string code = (string)(sel.Код ?? "");
                string name = (string)(sel.Наименование ?? "");
                Console.WriteLine($"  Code={code}  Name={name}");
                count++;
                if (count > 30) { Console.WriteLine("  ... (truncated)"); break; }
            }
            Marshal.ReleaseComObject(conn);
            Marshal.ReleaseComObject(connector);
            return;
        }

        if (command == "meta" && args.Length > 1)
        {
            // Dump document metadata (attribute names) for a given database
            string db = args[1];
            Console.OutputEncoding = Encoding.UTF8;
            var connectorType = Type.GetTypeFromProgID("V83.COMConnector")
                ?? throw new Exception("V83.COMConnector not registered");
            dynamic connector = Activator.CreateInstance(connectorType)!;
            var connStr = $"Srvr=\"{loaderConfig.Server}\";Ref=\"{db}\";" +
                          $"Usr=\"{loaderConfig.User}\";Pwd=\"{loaderConfig.Password}\"";
            dynamic conn = connector.Connect(connStr);
            foreach (var docType in new[] { "СписаниеСРасчетногоСчета", "ПоступлениеНаРасчетныйСчет" })
            {
                Console.WriteLine($"\n=== {docType} attributes ===");
                dynamic mgr = conn.Документы.GetType().InvokeMember(
                    docType, System.Reflection.BindingFlags.GetProperty,
                    null, (object)conn.Документы, null);
                dynamic newDoc = mgr.СоздатьДокумент();
                dynamic meta = newDoc.Метаданные();
                dynamic attrs = meta.Реквизиты;
                for (int i = 0; i < (int)attrs.Количество(); i++)
                {
                    dynamic attr = attrs.Получить(i);
                    Console.WriteLine($"  {attr.Имя}");
                }
            }
            Marshal.ReleaseComObject(conn);
            Marshal.ReleaseComObject(connector);
            return;
        }

        if (command == "scan")
        {
            // Just scan databases and save mapping
            using var loggerFactory = LoggerFactory.Create(b => b.AddConsole());
            var logger = loggerFactory.CreateLogger<Program>();
            var com = new Com1CConnector(loaderConfig,
                loggerFactory.CreateLogger("COM"));
            com.ScanDatabases();
            Console.WriteLine($"Mapping saved. {com.AccountMap.Count} accounts found.");
            return;
        }

        if (command == "test" && args.Length > 1)
        {
            // Test parse a file (dry run, no loading)
            var filePath = args[1];
            var parsed = BankFileParser.Parse(filePath);
            Console.WriteLine($"File: {filePath}");
            Console.WriteLine($"Account: {parsed.Account}");
            Console.WriteLine($"Period: {parsed.DateStart} - {parsed.DateEnd}");
            Console.WriteLine($"Documents: {parsed.Documents.Count}");
            foreach (var doc in parsed.Documents)
            {
                var dir = doc.IsDebit(parsed.Account) ? "DEBIT" : "CREDIT";
                Console.WriteLine($"  {dir} #{doc.Number} {doc.Date} " +
                    $"{doc.Amount:F2} {doc.Purpose[..Math.Min(60, doc.Purpose.Length)]}");
            }

            using var loggerFactory = LoggerFactory.Create(b => b.AddConsole());
            var com = new Com1CConnector(loaderConfig,
                loggerFactory.CreateLogger("COM"));
            com.LoadCachedMapping();
            var mapping = com.FindAccount(parsed.Account);
            if (mapping != null)
                Console.WriteLine($"\nMapped to: {mapping.Database} ({mapping.OrgName})");
            else
                Console.WriteLine($"\nWARNING: Account {parsed.Account} not found in mapping!");
            return;
        }

        if (command == "load" && args.Length > 1)
        {
            // Load a single file
            var filePath = args[1];
            using var loggerFactory = LoggerFactory.Create(b => b.AddConsole());
            var logger = loggerFactory.CreateLogger("COM");
            var com = new Com1CConnector(loaderConfig, logger);
            if (!com.LoadCachedMapping())
            {
                Console.WriteLine("No cached mapping. Run 'scan' first.");
                return;
            }
            var parsed = BankFileParser.Parse(filePath);
            var result = com.LoadFile(parsed);
            Console.WriteLine(result.Success
                ? $"OK: {result.Created} documents created"
                : $"FAILED: {result.Error}");
            return;
        }

        // Default: run as service
        var builder = Host.CreateApplicationBuilder(args);
        builder.Services.AddSingleton(loaderConfig);
        builder.Services.AddSingleton(sp =>
            new Com1CConnector(loaderConfig,
                sp.GetRequiredService<ILogger<Com1CConnector>>()));
        builder.Services.AddHostedService<LoaderWorker>();
        builder.Services.AddWindowsService(options =>
            options.ServiceName = "Loader1C");

        // Configure logging: console + file
        builder.Logging.ClearProviders();
        builder.Logging.AddConsole();
        if (!string.IsNullOrEmpty(loaderConfig.LogFile))
        {
            var logDir = Path.GetDirectoryName(loaderConfig.LogFile);
            if (!string.IsNullOrEmpty(logDir))
                Directory.CreateDirectory(logDir);
            builder.Logging.AddProvider(new FileLoggerProvider(loaderConfig.LogFile));
        }

        var host = builder.Build();
        host.Run();
    }
}
