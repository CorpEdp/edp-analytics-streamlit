# EDP Analytics Windows Desktop Wrapper

This project packages the hosted EDP Analytics Streamlit application in a native Windows window. It does not include or modify the Streamlit source code and requires an internet connection at runtime.

## Build

From this directory on Windows 10 or Windows 11:

```bat
build.bat
```

The script creates `.venv`, installs `requirements.txt`, and builds:

```text
dist\EDP Analytics.exe
```

The generated executable is a GUI application, so it does not open a command prompt window. The target computer does not need Python, VS Code, Streamlit, or PyInstaller installed.

## Test

After a successful build, launch the executable:

```bat
dist\EDP Analytics.exe
```

The wrapper checks the hosted URL before opening it. If the URL cannot be reached, it displays a readable connection error inside the application window.

The packaged application uses the WebView2 browser runtime supplied by Windows or Microsoft Edge. If a Windows installation does not have WebView2 available, install the Microsoft Edge WebView2 Runtime once on that machine.