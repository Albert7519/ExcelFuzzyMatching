# ExcelFuzzyMatching 模糊匹配工具

本项目是一个基于 Django 的 Web 应用，支持 Excel 文件的模糊匹配、标准化处理与高亮标记。适用于数据清洗、批量标准化等场景。

## 主要特性

- 支持 Excel 文件（.xlsx/.xls）上传、模糊匹配处理与结果下载
- 两种处理模式：自学习标准化、参照标准匹配
- 处理后自动高亮所有被标准化的单元格
- 支持预览匹配结果
- 上传文件不再本地保存，仅处理和下载时生成本地文件，轻量高效
- 处理结果文件名自动为“源文件名+_processed”

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 初始化数据库

由于项目使用了数据库模型来存储处理记录和学习模式，首次运行或模型更新后需要进行数据库迁移：

```bash
python manage.py migrate
```

### 3. 启动开发服务器

```bash
python manage.py runserver
```

### 4. 访问

在浏览器中打开 [http://127.0.0.1:8000/excel/](http://127.0.0.1:8000/excel/)

## 目录结构说明

- `excel_matcher/`：核心业务应用，包含上传、处理、预览、下载等功能
    - `views.py`：主要视图逻辑，处理文件上传、预览、处理、下载
    - `services/excel_service.py`：Excel 文件处理与模糊匹配算法实现
    - `models.py`：处理记录与模式存储模型
    - `static/`、`templates/`：前端静态资源与页面模板
- `hello/`：示例应用（可选）
- `media/processed/`：仅用于存放处理后的 Excel 文件，上传文件不再保存
- `web_django/`：Django 项目配置
- `requirements.txt`：依赖包列表
- `manage.py`：Django 管理脚本

## 使用说明

1. 上传 Excel 文件（.xlsx/.xls），文件不会在本地保存，仅用于后续处理
2. 选择处理模式：
    - **参照标准匹配模式**：选择一列作为标准，其他列与其进行模糊匹配
    - **自学习标准化模式**：系统自动学习每列的数据规律进行标准化
3. 选择需要处理的列和匹配阈值，可预览标准化效果
4. 点击“处理文件”后，系统生成处理结果文件，所有被标准化的单元格会自动高亮
5. 点击“下载结果文件”即可获取，文件名为“源文件名+_processed”

## 依赖

- Django >= 5.0
- pandas
- openpyxl
- rapidfuzz

## 注意事项

- 上传的 Excel 文件不会在本地保存，仅在处理和下载时生成临时文件

## 扩展建议

- 可集成 Django REST Framework 实现 API 化
- 支持多用户隔离与权限管理
- 增加历史记录、任务队列、异步处理等高级功能

## 联系与反馈

如有建议或问题，欢迎提交 issue 或联系开发者。
