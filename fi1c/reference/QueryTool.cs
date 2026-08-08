using System;
using System.Runtime.InteropServices;

class QueryTool
{
    static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        string db = args.Length > 0 ? args[0] : "base_m";

        var connectorType = Type.GetTypeFromProgID("V83.COMConnector");
        dynamic connector = Activator.CreateInstance(connectorType!)!;
        dynamic conn = connector.Connect($"Srvr=\"navus-server\";Ref=\"{db}\";Usr=\"Zaykov Andrey\";Pwd=\"19700214\"");
        Console.WriteLine($"=== DATABASE: {db} ===");

        Run((object)conn, "МСФО records by year",
            @"ВЫБРАТЬ ГОД(Р.Период) КАК Г, КОЛИЧЕСТВО(*) КАК Н
              ИЗ РегистрБухгалтерии.МСФО КАК Р
              СГРУППИРОВАТЬ ПО ГОД(Р.Период)
              УПОРЯДОЧИТЬ ПО Г",
            sel => $"  {sel.Г}: {sel.Н}");

        Run((object)conn, "Хозрасчетный records by year",
            @"ВЫБРАТЬ ГОД(Р.Период) КАК Г, КОЛИЧЕСТВО(*) КАК Н
              ИЗ РегистрБухгалтерии.Хозрасчетный КАК Р
              СГРУППИРОВАТЬ ПО ГОД(Р.Период)
              УПОРЯДОЧИТЬ ПО Г",
            sel => $"  {sel.Г}: {sel.Н}");

        Run((object)conn, "МСФО 2025 by organization",
            @"ВЫБРАТЬ Р.Организация.Наименование КАК Орг, КОЛИЧЕСТВО(*) КАК Н
              ИЗ РегистрБухгалтерии.МСФО КАК Р
              ГДЕ ГОД(Р.Период) = 2025
              СГРУППИРОВАТЬ ПО Р.Организация.Наименование
              УПОРЯДОЧИТЬ ПО Н УБЫВ",
            sel => $"  {sel.Орг}: {sel.Н}");

        Console.WriteLine("=== DONE ===");
        Marshal.FinalReleaseComObject(conn);
    }

    static void Run(object connObj, string title, string text, Func<dynamic, string> fmt)
    {
        Console.WriteLine($"--- {title} ---");
        try
        {
            dynamic q = ((dynamic)connObj).NewObject("Запрос");
            q.Текст = text;
            dynamic sel = q.Выполнить().Выбрать();
            int rows = 0;
            while (sel.Следующий() && rows < 60) { Console.WriteLine(fmt(sel)); rows++; }
            if (rows == 0) Console.WriteLine("  (no rows)");
        }
        catch (Exception e) { Console.WriteLine($"  FAILED: {e.Message.Split('\n')[0]}"); }
    }
}
