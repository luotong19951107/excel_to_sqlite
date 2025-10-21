# Excel to SQLite Converter

Excel到SQLite转换工具，支持批量转换多个Excel文件。

## 使用方法

1. 将Excel文件放入 `input` 目录
2. 运行转换脚本：
   ```bash
   python excel_to_sqlite.py
   ```
3. 在 `output` 目录查看生成的 `.db` 文件

## 查看数据库

使用在线SQLite查看器：

| 网站                                                                            | 说明              |
| ----------------------------------------------------------------------------- | --------------- |
| [SQLite Editor (sqliteviewer.app)](https://sqliteviewer.app/)                 | UI 简洁，支持导出      |

**使用方法**：
1. 打开 [SQLite Editor](https://sqliteviewer.app/)
2. 将 `output` 目录中的 `.db` 文件拖拽到网页中
3. 即可查看数据库内容、表结构、执行SQL查询等

## 安装依赖

```bash
pip install -r requirements.txt
```
