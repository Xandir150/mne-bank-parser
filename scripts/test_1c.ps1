chcp 65001 | Out-Null
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$connector = New-Object -ComObject "V83.COMConnector"
$conn = $connector.Connect('Srvr="navus-server";Ref="base_m";Usr="Zaykov Andrey";Pwd="19700214"')
Write-Host "Connected to base_m"

# Try to get metadata catalog names via XML
$md = $conn.Metadata
Write-Host "Metadata name: $($md.Name)"
Write-Host "Metadata synonym: $($md.Synonym)"

# Try direct access to well-known objects
try {
    $orgs = $conn.Catalogs
    Write-Host "Catalogs type: $($orgs.GetType())"
} catch {
    Write-Host "Catalogs failed: $_"
}

# Try Russian property names via InvokeMember
try {
    $type = $conn.GetType()
    $catalogs = $type.InvokeMember(
        [char]0x0421 + [char]0x043F + [char]0x0440 + [char]0x0430 + [char]0x0432 + [char]0x043E + [char]0x0447 + [char]0x043D + [char]0x0438 + [char]0x043A + [char]0x0438,
        [System.Reflection.BindingFlags]::GetProperty,
        $null, $conn, $null)
    Write-Host "Got catalogs via InvokeMember"

    $orgCat = $type.InvokeMember(
        [char]0x0421 + [char]0x043F + [char]0x0440 + [char]0x0430 + [char]0x0432 + [char]0x043E + [char]0x0447 + [char]0x043D + [char]0x0438 + [char]0x043A + [char]0x0438,
        [System.Reflection.BindingFlags]::GetProperty,
        $null, $conn, $null)
    Write-Host "OrgCat: $orgCat"
} catch {
    Write-Host "InvokeMember failed: $_"
}

# Simplest test - just eval an expression
try {
    $testStr = $conn.String("test123")
    Write-Host "String test: $testStr"

    $evalResult = $conn.EvalExpression("1+1")
    Write-Host "Eval 1+1: $evalResult"
} catch {
    Write-Host "Eval failed: $_"
}
