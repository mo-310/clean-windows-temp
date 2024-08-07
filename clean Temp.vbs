' Create a shell object
Set shell = CreateObject("WScript.Shell")

' Function to clear temp directories
Sub ClearTempFolder(folderPath)
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then
        Set folder = fso.GetFolder(folderPath)
        For Each file In folder.Files
            file.Delete True
        Next
        For Each subFolder In folder.SubFolders
            subFolder.Delete True
        Next
    End If
    On Error GoTo 0
End Sub

' Clear %TEMP% folder
tempPath = shell.ExpandEnvironmentStrings("%TEMP%")
ClearTempFolder tempPath

' Clear %WINDIR%\Temp folder
windirTempPath = shell.ExpandEnvironmentStrings("%WINDIR%\Temp")
ClearTempFolder windirTempPath

' Clear Prefetch folder
prefetchPath = shell.ExpandEnvironmentStrings("%WINDIR%\Prefetch")
ClearTempFolder prefetchPath

WScript.Echo "Temporary files and folders have been cleared."
