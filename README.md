# 批量内容与文件名替换工具

## 功能
- 支持批量替换文本/超文本（.txt/.html/.htm）、Word（.doc/.docx）、Excel（.xls/.xlsx）、PowerPoint（.ppt/.pptx）等文件内容
- 支持文件名同步替换
- Word文档通配符、带格式、半角全角转换
- PowerPoint多母版替换
- Excel批量替换
- 智能编码识别、Unicode支持
- 进度条与日志输出

## 使用方法
1. 安装依赖  
   `pip install -r requirements.txt`
2. 运行  
   `python main.py`
3. 打包为exe  
   `pyinstaller --onefile --noconsole main.py` 