
使用方法
1. 克隆项目
   git clone https://github.com/zhkali127/vuln-report-extractor.git
   cd vuln-report-extractor

2. 安装依赖
   pip install -r requirements.txt

3. 将扫描器生成的报告文件放入 `input/` 目录

4. 运行脚本
   python main.py

5. 转换完成后，可以在 `output/` 目录下找到生成的 Excel 文件

目录结构示例
- input/                存放扫描器报告 (XML/HTML/CSV/JSON)
- output/               存放最终结果 Excel 文件
- macros/               存放 Nessus/绿盟 宏文件 (xlsm)
- scripts/              各类处理脚本
- main.py               主入口

作者信息
- 作者: zhkali127
- GitHub: https://github.com/zhkali127
