using System;
using System.Reflection;
using System.Runtime.InteropServices;

// Тест-раннер FinIskaziCG.epf: подключает обработку через временное хранилище,
// заполняет по организации/году, выгружает XML/XLSX/PDF/JSON.
// Usage: RunEpf.exe <db> <epfPath> <orgName> <year> <outDir>
class RunEpf
{
    static object Call(object obj, string method, params object?[] a)
    {
        try
        {
            return obj.GetType().InvokeMember(method, BindingFlags.InvokeMethod, null, obj, a)!;
        }
        catch (TargetInvocationException e)
        {
            Console.WriteLine($"1C ERROR in {method}: {e.InnerException?.Message}");
            throw;
        }
    }

    static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        string db = args[0], epf = args[1], org = args[2], outDir = args[4];
        int year = int.Parse(args[3]);

        var t = Type.GetTypeFromProgID("V83.COMConnector");
        dynamic connector = Activator.CreateInstance(t!)!;
        dynamic conn = connector.Connect($"Srvr=\"navus-server\";Ref=\"{db}\";Usr=\"Zaykov Andrey\";Pwd=\"19700214\"");
        Console.WriteLine($"connected to {db}");

        dynamic prot = conn.NewObject("ОписаниеЗащитыОтОпасныхДействий");
        prot.ПредупреждатьОбОпасныхДействиях = false;
        dynamic bd = conn.NewObject("ДвоичныеДанные", epf);
        string addr = conn.ПоместитьВоВременноеХранилище(bd);
        object ext = conn.ВнешниеОбработки;
        object name = Call(ext, "Подключить", addr, "FinIskaziCG", false, prot);
        Console.WriteLine($"attached: {name}");
        object obj = Call(ext, "Создать", name, false);

        Call(obj, "ЗаполнитьПоПараметрам", org, year);
        Console.WriteLine("filled");

        string ctrl = (string)Call(obj, "ПроверитьКонтроли");
        Console.WriteLine(ctrl.Length == 0 ? "controls: OK" : $"controls:\n{ctrl}");

        Console.WriteLine(Call(obj, "ВыгрузитьXMLФИ", $@"{outDir}\fi_{year}.xml"));
        Console.WriteLine(Call(obj, "ВыгрузитьXMLПД", $@"{outDir}\pd_{year}.xml"));
        Console.WriteLine(Call(obj, "ВыгрузитьОтчет", $@"{outDir}\fi_{year}.xlsx", "XLSX"));
        Console.WriteLine(Call(obj, "ВыгрузитьОтчет", $@"{outDir}\fi_{year}.pdf", "PDF"));

        string json = (string)Call(obj, "ПолучитьСтрокиJSON");
        System.IO.File.WriteAllText($@"{outDir}\rows_{year}.json", json, new System.Text.UTF8Encoding(false));
        Console.WriteLine($"rows json: {json.Length} chars");
        Console.WriteLine("DONE");
        Marshal.FinalReleaseComObject(conn);
    }
}
