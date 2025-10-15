@echo off
chcp 65001
echo ========================================
echo 自动化邮件服务系统 - Windows打包脚本
echo ========================================
echo.

echo [1/3] 检查Python环境...
python --version
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python 3.7+
    pause
    exit /b 1
)
echo.

echo [2/3] 安装依赖包...
pip install -r requirements.txt
if errorlevel 1 (
    echo 错误: 依赖包安装失败
    pause
    exit /b 1
)
echo.

echo [3/3] 打包程序...
pyinstaller build.spec --clean
if errorlevel 1 (
    echo 错误: 打包失败
    pause
    exit /b 1
)
echo.

echo ========================================
echo 打包完成！
echo 可执行文件位置: dist\EmailService.exe
echo ========================================
pause
