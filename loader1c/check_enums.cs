using System;
using System.Runtime.InteropServices;

class CheckEnums
{
    static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        var connectorType = Type.GetTypeFromProgID("V83.COMConnector");
        dynamic connector = Activator.CreateInstance(connectorType);
        dynamic conn = connector.Connect(
            "Srvr=\"navus-server\";Ref=\"lardaux\";Usr=\"Zaykov Andrey\";Pwd=\"19700214\"");

        // Find unposted zajm doc, set fields, post — all with Загрузка
        Console.WriteLine("=== Full cycle: set + post with Загрузка ===");
        try
        {
            dynamic q = conn.NewObject("Запрос");
            q.Текст = @"ВЫБРАТЬ ПЕРВЫЕ 1 Док.Ссылка, Док.Номер
            ИЗ Документ.ПоступлениеНаРасчетныйСчет КАК Док
            ГДЕ Док.НазначениеПлатежа ПОДОБНО ""%zajm%""
                И Док.Дата МЕЖДУ &Д1 И &Д2
                И НЕ Док.ПометкаУдаления И НЕ Док.Проведен
            УПОРЯДОЧИТЬ ПО Док.Номер УБЫВ";
            q.УстановитьПараметр("Д1", new DateTime(2026, 3, 19));
            q.УстановитьПараметр("Д2", new DateTime(2026, 3, 20));
            dynamic res = q.Выполнить().Выбрать();
            if (!res.Следующий()) { Console.WriteLine("  No unposted zajm docs"); return; }

            Console.WriteLine($"  Doc: {res.Номер}");
            dynamic docRef = res.Ссылка;

            // Step 1: ensure fields are set
            dynamic obj = docRef.ПолучитьОбъект();
            dynamic ddsRef = conn.Справочники.СтатьиДвиженияДенежныхСредств
                .НайтиПоНаименованию("Получение кредитов и займов", true);
            dynamic acctRef = conn.ПланыСчетов.Хозрасчетный.НайтиПоКоду("67.03");

            obj.СтатьяДвиженияДенежныхСредств = ddsRef;
            obj.СчетУчетаРасчетовСКонтрагентом = acctRef;
            obj.ОбменДанными.Загрузка = true;
            obj.Записать();
            Console.WriteLine("  Fields set and saved");

            // Step 2: post
            dynamic obj2 = docRef.ПолучитьОбъект();

            // Verify fields before posting
            bool ddsOk = !(bool)obj2.СтатьяДвиженияДенежныхСредств.Пустая();
            bool acctOk = !(bool)obj2.СчетУчетаРасчетовСКонтрагентом.Пустая();
            Console.WriteLine($"  Before post: ДДС={ddsOk} Счет={acctOk}");

            obj2.ОбменДанными.Загрузка = true;
            obj2.Записать(conn.РежимЗаписиДокумента.Проведение);
            Console.WriteLine("  POSTED OK!");

            // Verify posted
            dynamic check = docRef.ПолучитьОбъект();
            Console.WriteLine($"  Проведен: {check.Проведен}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Error: {ex.Message}");
            if (ex.InnerException != null)
                Console.WriteLine($"  Inner: {ex.InnerException.Message}");
        }

        Marshal.ReleaseComObject(conn);
        Marshal.ReleaseComObject(connector);
    }
}
