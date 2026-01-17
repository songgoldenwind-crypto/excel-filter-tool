@echo off
echo ========================================
echo Excel筛选工具打包脚本
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python 3.8或更高版本
    pause
    exit /b 1
)

echo 正在安装依赖...
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo 依赖安装失败
    pause
    exit /b 1
)

echo.
echo 正在打包程序...
pyinstaller --onefile --windowed --name="Excel数据筛选工具" --icon=NONE --clean excel_filter_tool.py
if errorlevel 1 (
    echo 打包失败
    pause
    exit /b 1
)

echo.
echo ========================================
echo 打包完成！
echo 可执行文件位置: dist\Excel数据筛选工具.exe
echo ========================================
echo.
pause
