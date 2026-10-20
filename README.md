# ExcelKiller Excel填补助手

## 构建后，请使用 “管理员” 运行 exe

> **此工具是本人为了帮哥们儿简化政府工作内容写的工具**
> 
> 如果各位觉得有什么修改的地方，可以Fork本仓库自己改改

## 功能简要说明

把一堆格式各不相同的 Excel 表格，按“身份证/某个键”自动对起来补全、或者把各村台账汇总成一张县级总表，省掉人工一列列复制粘贴的活。

## 功能页

### ① 信息补全

按匹配键（身份证/姓名+手机/自定义多列）把源表的字段填进目标表。

- 目标表工作表可**多选**，自动与源表工作表按名字对应（可手动改）
- 匹配键**自动识别**（身份证 → 姓名+手机 → 手机 → 姓名 → 同名唯一列）
- 字段映射**自动匹配**（含同义列名），可逐个改成 默认/仅填空/覆盖/跳过
- 写入方式：仅填空 / 覆盖 / 仅新增 / 填空并新增 / 覆盖并新增
- **防错行**：键对上但姓名对不上时不写入，并在明细清单单独列出
- **源表冲突**：多源表给同一格不同值时弹窗让你选，选择可记住，也能一键清除记忆
- 只处理可见行开关（表格被筛选时）
- 先预演、后执行；默认另存新文件，原文件不动

### ② 台账合并

- 源表支持**多选文件**，内部按工作表名自动对上
- 源工作表清单默认全选，可按住 Ctrl 自选；且**跟随目标表勾选联动**
- 空白继承上一行（同户多人）、镇/村名自动统一、序号自动重排
- 跳过模板底部预格式化空行，从表头下第一行开始写
- 输出同时生成“合并明细”，记录每个表实际追加了哪些行

## 运行

需要 Python 3.9+ 和 `openpyxl`，界面用标准库 `tkinter`。

```bash
pip install openpyxl
python excel_tool.py
```

打包成免安装 exe：

> PyInstaller

```bash
python -m PyInstaller --noconfirm --onedir --windowed --name "ExcelKiller" ^
  --exclude-module numpy --exclude-module pandas --exclude-module matplotlib ^
  --exclude-module PIL --exclude-module scipy --exclude-module IPython ^
  excel_tool.py
```

> 用 `--onedir` 而不是 `--onefile`：单文件版会往临时目录释放文件，容易被安全软件拦。

## 目录说明

| 文件 | 说明 |
|---|---|
| `excel_tool.py` | 主程序（引擎 + 界面，单文件） |
| `make_rule.py` | 生成示例规则 json |
| `smoke_test.py` 等 `test_*.py` | 开发期测试脚本（需自备样例表格，按脚本内路径调整） |
