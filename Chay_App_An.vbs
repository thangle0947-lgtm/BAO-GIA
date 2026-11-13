Set WshShell = CreateObject("WScript.Shell")

' 1. Khởi chạy Streamlit ngầm (tắt tự động mở tab trình duyệt mặc định)
WshShell.Run "cmd /c py -m streamlit run app.py --server.headless true", 0, False

' 2. Cho server Streamlit 2 giây để khởi động
WScript.Sleep 2000

' 3. Mở cửa sổ ứng dụng riêng biệt (App Mode) bằng Microsoft Edge
WshShell.Run "msedge --app=http://localhost:8501", 1, False