using System;
using System.Reflection;
using System.Runtime.InteropServices;
class AttachTest
{
    static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        string db = args[0]; string path = args[1];
        var t = Type.GetTypeFromProgID("V83.COMConnector");
        dynamic connector = Activator.CreateInstance(t!)!;
        dynamic conn = connector.Connect($"Srvr=\"navus-server\";Ref=\"{db}\";Usr=\"Zaykov Andrey\";Pwd=\"19700214\"");
        Console.WriteLine("connected");
        dynamic prot = conn.NewObject("ОписаниеЗащитыОтОпасныхДействий");
        prot.ПредупреждатьОбОпасныхДействиях = false;
        dynamic bd = conn.NewObject("ДвоичныеДанные", path);
        string addr = conn.ПоместитьВоВременноеХранилище(bd);
        Console.WriteLine($"tmpstore: {addr}");
        object ext = conn.ВнешниеОбработки;
        try
        {
            object name = ext.GetType().InvokeMember("Подключить", BindingFlags.InvokeMethod, null, ext,
                new object?[] { addr, "SmokeTest", false, prot });
            Console.WriteLine($"attached as: {name}");
            object obj = ext.GetType().InvokeMember("Создать", BindingFlags.InvokeMethod, null, ext,
                new object?[] { name, false })!;
            object pong = obj.GetType().InvokeMember("Пинг", BindingFlags.InvokeMethod, null, obj, new object?[] { })!;
            Console.WriteLine($"ping: {pong}");
        }
        catch (TargetInvocationException e)
        {
            var inner = e.InnerException;
            Console.WriteLine($"1C ERROR: {inner?.GetType().Name}: {inner?.Message}");
            if (inner is COMException ce) Console.WriteLine($"HRESULT: 0x{ce.HResult:X8}");
        }
        Marshal.FinalReleaseComObject(conn);
    }
}
