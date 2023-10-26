@echo off
chcp 65001
echo 正在激活虚拟环境...
call ..\venv\Scripts\activate

echo 清理 package 目录...
if exist package rmdir /s /q package

echo 创建 package 目录...
mkdir package

echo 正在使用 PyInstaller 打包 main.py...
pyinstaller -y --distpath .\package ..\main.py

echo 删除临时 build 文件夹...
rmdir /s /q .\build

echo 复制三方库和配置文件...
mkdir package\main\_internal\tabula
copy ..\venv\Lib\site-packages\tabula\tabula-1.0.5-jar-with-dependencies.jar package\main\_internal\tabula\tabula-1.0.5-jar-with-dependencies.jar
copy ..\config.ini package\main\config.ini

echo 打包完成！
pause

echo 正在退出虚拟环境...
deactivate

echo 脚本执行完毕。
