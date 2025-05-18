' VBS Launcher Maker & Deleter (OG Version)

Dim fso, shell, desktop, choice
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
desktop = shell.SpecialFolders("Desktop")

choice = LCase(InputBox("Type 'make' to create a shortcut or 'delete' to remove one:"))

If choice = "make" Then
    Dim url, name, filename, scriptPath, launcher
    url = InputBox("Enter the URL:")

    If Left(url, 4) <> "http" Then
        MsgBox "🚫 Invalid URL. Must start with http or https."
        WScript.Quit
    End If

    name = url
    name = Replace(name, "https://", "")
    name = Replace(name, "http://", "")
    name = Replace(name, "www.", "")
    name = Split(name, "/")(0)
    filename = Replace(name, ".", "_") & ".vbs"
    scriptPath = desktop & "\" & filename

    launcher = "Set shell = CreateObject(""WScript.Shell"")" & vbCrLf & _
               "shell.Run """ & url & """"

    Set file = fso.CreateTextFile(scriptPath, True)
    file.WriteLine launcher
    file.Close

    MsgBox "✅ Launcher created on Desktop: " & filename

ElseIf choice = "delete" Then
    Dim delName, delPath
    delName = InputBox("Enter the name of the shortcut (without .vbs is fine):")
    If Right(delName, 4) <> ".vbs" Then
        delName = delName & ".vbs"
    End If
    delPath = desktop & "\" & delName

    If fso.FileExists(delPath) Then
        fso.DeleteFile delPath
        MsgBox "🗑️ Deleted from Desktop: " & delName
    Else
        MsgBox "❌ Not found on Desktop: " & delName
    End If

Else
    MsgBox "🤔 Invalid choice. Type 'make' or 'delete'."
End If
