'Script path
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

Set file = fso.OpenTextFile(scriptdir + "\test.txt", 1)
Set outputFile = fso.CreateTextFile(scriptdir + "\output.txt")

'This function return True if it gets reply
'computerId : specified computer name or ip adress for pinging
function Ping( computerId )

    Dim colPingResults, objPingResult, strQuery

    ' Define the WMI query
    strQuery = "SELECT * FROM Win32_PingStatus WHERE Address = '" & computerId & "'"

    ' Run the WMI query
    Set colPingResults = GetObject("winmgmts://./root/cimv2").ExecQuery( strQuery )

    for each item in colPingResults
        if not IsObject ( item ) Then
            Ping = false
        elseif item.StatusCode = 0 Then
            Ping = True
        else
            Ping = false
        end if
    next
    
    Set colPingResults = Nothing

end function

Do until file.AtEndOfStream = True
    computerId = file.ReadLine
    if  Ping(computerId) Then
        outputFile.WriteLine(computerId)
    end if
loop