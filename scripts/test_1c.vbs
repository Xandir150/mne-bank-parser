Set connector = CreateObject("V83.COMConnector")
Set conn = connector.Connect("Srvr=""navus-server"";Ref=""base_m"";Usr=""Zaykov Andrey"";Pwd=""19700214""")

WScript.Echo "=== Connected to base_m ==="

WScript.Echo "Organizations:"
Set orgSel = conn.Справочники.Организации.Выбрать()
Do While orgSel.Следующий()
    WScript.Echo "  " & orgSel.Наименование & " | INN: " & orgSel.ИНН
Loop

WScript.Echo "Bank accounts:"
Set acctSel = conn.Справочники.БанковскиеСчета.Выбрать()
Do While acctSel.Следующий()
    WScript.Echo "  " & acctSel.НомерСчета & " | " & acctSel.Владелец.Наименование
Loop
