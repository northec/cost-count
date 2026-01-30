# 文件积分统计工具

用于统计目录中的 PDF、DGN、DWG 文件，计算积分并生成 Excel 报告。
可直接运行脚本，也可使用打包后的 EXE。

## 功能

- 扫描指定目录（含子目录）中的 PDF、DGN、DWG 文件
- 根据实际页面尺寸自动换算成 A4 页数
- 按照规则计算积分
- 生成带时间戳的 Excel 统计报告
- 自动打开生成的报告
- 支持通过环境变量关闭自动打开

## 积分规则

### PDF 文件
- 10 积分起步（含 3 页 A4）
- 超出 3 页，每页 3 积分
- 大于 A4 的页面按面积比换算成 A4 页数

### DGN/DWG 文件
- 按 16 页 PDF 计算，固定 49 积分/文件

## 安装依赖

```bash
pip install openpyxl pypdf
```

## 使用方法

```bash
# 统计当前目录
python file_counter.py

# 统计指定目录
python file_counter.py /path/to/directory
```

## EXE 使用方法

已打包的 EXE 位于 `dist/` 目录，命名规则为 `file_counter_YYYY.MM.DD.N.exe`。

```bash
# 统计当前目录
dist/file_counter_2026.01.30.1.exe

# 统计指定目录
dist/file_counter_2026.01.30.1.exe "D:\path\to\directory"
```

如不希望自动打开 Excel，可设置环境变量：

```bash
set FILE_COUNTER_NO_OPEN=1
```

更完整的 EXE 使用说明请见 `使用说明.md`。

## 版本说明

版本说明请见 `更新记录.md`。

## 输出字段

| 字段 | 说明 |
|------|------|
| 路径 | 文件完整路径 |
| 文件名 | 文件名称 |
| 类型 | PDF/DGN/DWG |
| 大小 | 文件大小（K/M） |
| 页数 | 实际页数 |
| A4页数 | 换算成A4的页数 |
| 积分 | 该文件积分 |

## 统计数据

报告包含以下汇总信息：
- 各类型文件个数、总体积、总页数、总积分
- 总积分量

## 示例

```bash
$ python file_counter.py ./project
正在扫描目录: /Users/xxx/project
找到 25 个文件
报告已生成: /Users/xxx/project/文件统计报告_20260129_120000.xlsx
```
