@echo off
REM 切换控制台编码为 UTF-8，支持中文输出
chcp 65001

REM 遍历当前文件夹所有 py 文件
for %%f in (*.py) do (
    echo 正在运行 %%f ...
    python "%%f"
    echo ----------------------------
)

REM 执行完成，保持窗口打开
echo 所有 Python 文件已执行完成.
pause
