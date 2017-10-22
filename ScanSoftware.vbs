Option Explicit
Const HKEY_LOCAL_MACHINE = &H80000002
Dim oReg : Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Dim oShell : Set oShell = CreateObject("WScript.Shell")
Dim oNetwork : Set oNetwork = CreateObject("WScript.Network")
Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")
Dim oFile
Dim sPath, aSub, sKey, sSearch
Dim szValue
Dim sRes
Dim sTmpRes
Dim searchFor : Set searchFor = CreateObject("System.Collections.ArrayList")
Dim saveToPath
 
searchFor.Add("Microsoft Office")
saveToPath = "C:\Users\ladmin\Desktop"
 
' Get all keys within sPath
sPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
oReg.EnumKey HKEY_LOCAL_MACHINE, sPath, aSub
 
' Loop through each key
For Each sKey In aSub
    szValue = Null
    oReg.GetStringValue HKEY_LOCAL_MACHINE, sPath & "\" & sKey, "DisplayName", szValue
    if (not IsNull(szValue)) and (not IsEmpty(szValue)) then
        For Each sSearch In searchFor
            if InStr(1, LCase(szValue), LCase(sSearch)) = 1 then
                sRes = sRes & vbNewLine & szValue
            end if
        Next
    end if
Next
 
if IsNull(sRes) or IsEmpty(sRes) then
    WScript.Echo "Nothing found."
else
    For Each sSearch In searchFor
        sTmpRes = sTmpRes & sSearch & vbNewLine
    Next
    sRes = "Software found on '" & oNetwork.ComputerName & "' based on search terms:" & vbNewLine & sTmpRes & vbNewLine & "Result:" & sRes
    Set oFile = oFso.CreateTextFile(saveToPath & "\" & oNetwork.ComputerName & ".txt",True)
    oFile.Write(sRes & vbCrLf)
    oFile.Close
    WScript.Echo sRes & vbNewLine & vbNewLine & "Saved to file: " & saveToPath & "\" & oNetwork.ComputerName & ".txt"
end if
