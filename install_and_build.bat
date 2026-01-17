@echo off
chcp 65001 >nul
echo ========================================
echo   Excel数据筛选工具 - 自动打包
echo ========================================
echo.

REM 检查Python是否已安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [1/4] 正在安装Python...
    echo 请在浏览器中下载并安装Python 3.10
    echo.
    echo 下载地址会自动打开...
    start https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe
    echo.
    echo 安装时请勾选 "Add Python to PATH"
    echo.
    pause

    REM 再次检查
    python --version >nul 2>&1
    if errorlevel 1 (
        echo Python未正确安装，请重试
        pause
        exit /b 1
    )
)

echo [2/4] Python已安装:
python --version
echo.

echo [3/4] 正在安装依赖包...
python -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple
python -m pip install openpyxl pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
echo.

echo [4/4] 正在打包程序...
echo 这可能需要1-2分钟，请耐心等待...
echo.

pyinstaller --onefile --windowed --name="Excel数据筛选工具" --clean excel_filter_tool.py

if exist "dist\Excel数据筛选工具.exe" (
    echo.
    echo ========================================
    echo   ✓ 打包成功！
    echo ========================================
    echo.
    echo 文件位置: %cd%\dist\Excel数据筛选工具.exe
    dir "dist\Excel数据筛选工具.exe" | find "Excel数据筛选工具.exe"
    echo.
    echo 现在可以将exe文件复制到任何Windows电脑使用了！
    echo.
    start dist
) else (
    echo.
    echo ========================================
    echo   ✗ 打包失败
    echo ========================================
    echo.
    echo 请检查错误信息并重试
)

echo.
pause
