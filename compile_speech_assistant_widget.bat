rem call c:\virtual\speechTotext\Scripts\activate

rem c:\virtual\speechTotext\Scripts\pyinstaller --onefile --windowed --icon=app_icon.ico --add-data="rec.gif;resources" --add-data="app_icon.ico;resources" --add-data="credentials.json;resources" speech_assistant_widget.py

rem C:\Python\Python27\Scripts\pyinstaller --onefile --icon=app_icon.ico --add-data="rec.gif;resources" --add-data="app_icon.ico;resources" --add-data="credentials.json;resources"  --hidden-import="google-cloud-speech" --hidden-import="google-cloud-core" --hidden-import="google-api-core" speech_assistant_widget.py

rem call e:\virtual\speech2\Scripts\activate
pyinstaller --onefile --windowed --icon=app_icon.ico --add-data="rec.gif;resources" --add-data="app_icon.ico;resources" --add-data="credentials.json;resources" --add-data="grpc.pem;resources" speech_assistant_widget.py
