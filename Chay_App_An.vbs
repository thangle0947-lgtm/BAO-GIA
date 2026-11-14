Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = folder

' 0. Doc khoa API tu file api_key.txt (dong dau tien) cho tien trinh con
keyFile = folder & "\api_key.txt"
If fso.FileExists(keyFile) Then
    Set f = fso.OpenTextFile(keyFile, 1)
    If Not f.AtEndOfStream Then
        WshShell.Environment("Process")("ANTHROPIC_API_KEY") = Trim(f.ReadLine())
    End If
    f.Close
End If

' 1. Khoi chay Streamlit ngam (tat tu dong mo tab trinh duyet mac dinh)
WshShell.Run "cmd /c py -m streamlit run app.py --server.headless true", 0, False

' 2. Cho server Streamlit 2 giay de khoi dong
WScript.Sleep 2000

' 3. Mo cua so ung dung rieng biet (App Mode) bang Microsoft Edge
WshShell.Run "msedge --app=http://localhost:8501", 1, False
