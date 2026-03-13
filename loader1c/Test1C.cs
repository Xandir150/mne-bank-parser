using System;
using System.Runtime.InteropServices;

class Test1C
{
    static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        var connectorType = Type.GetTypeFromProgID("V83.COMConnector");
        if (connectorType == null)
        {
            Console.WriteLine("ERROR: V83.COMConnector not registered");
            return;
        }

        dynamic connector = Activator.CreateInstance(connectorType);
        Console.WriteLine("COM connector created");

        dynamic conn = connector.Connect(
            "Srvr=\"navus-server\";Ref=\"base_m\";Usr=\"Zaykov Andrey\";Pwd=\"19700214\"");
        Console.WriteLine("Connected to base_m");

        // List organizations
        Console.WriteLine("\nOrganizations:");
        dynamic orgSel = conn.Справочники.Организации.Выбрать();
        while (orgSel.Следующий())
        {
            string name = orgSel.Наименование;
            string inn = orgSel.ИНН;
            Console.WriteLine($"  {name} | INN: {inn}");
        }

        // List bank accounts
        Console.WriteLine("\nBank accounts:");
        dynamic acctSel = conn.Справочники.БанковскиеСчета.Выбрать();
        while (acctSel.Следующий())
        {
            string acctNo = acctSel.НомерСчета;
            string orgName = acctSel.Владелец.Наименование;
            Console.WriteLine($"  {acctNo} | {orgName}");
        }

        Marshal.ReleaseComObject(conn);
        Marshal.ReleaseComObject(connector);
    }
}
