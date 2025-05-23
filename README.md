# Excel2Json

一个简单高效的 Excel 转 JSON 工具，支持单文件转换和批量目录处理。

## 功能特点

- 支持单个 Excel 文件转换和批量目录处理
- 自动检测文件变更，只处理发生变化的文件
- 支持自定义起始行号（默认从第2行开始读取数据）
- 提供详细的错误和警告信息
- 支持元数据管理，记录文件处理状态
- 自动创建输出目录结构
- 支持自动解析单元格内容格式，将数值、json格式以本身进行存储而非字符串格式

## 系统要求

- .NET Framework 4.7.2 或更高版本

## 使用方法

1. 下载最新版本的发布包
2. 运行程序，支持以下两种模式：
   - 单文件模式：处理单个 Excel 文件
   - 目录模式：批量处理指定目录下的所有 Excel 文件

## 命令行参数

```
-i, --input <path>    指定输入的 Excel 文件或文件夹路径（默认为当前目录下的 excel 文件夹）
-o, --output <path>   指定输出的 JSON 文件或文件夹路径（默认为当前目录下的 json 文件夹）
-m, --meta <path>     指定元数据文件路径（默认为当前目录下的 .e2jmeta 文件）
-r, --row <number>    指定开始读取数据的行号（默认为 2）
```

## 输出格式

程序会将 Excel 数据转换为以下 JSON 格式：
```json
{
    "data": [
        {
            "列名1": "值1",
            "列名2": "值2",
            ...
        },
        ...
    ]
}
```

### 自动解析单元格内容格式

程序支持自动解析单元格内容格式，将数值、json格式以本身进行存储而非字符串格式。以下是一个示例：

```json
{
    "data": [
        {
            "列名1": 123,
            "列名2": 123.456,
            "列名3": {"key": "value"},
            "列名4": [0, 1, 2, 3, 4],
            "列名5": "abcabcabc",
            ...
        },
        ...
    ]
}
```

## 注意事项

- Excel 文件的第一行将被视为列名
- 程序会自动跳过空行
- 支持增量更新，只处理发生变化的文件
- 元数据文件用于记录文件处理状态，请勿手动修改

## 许可证

本项目采用 GNU General Public License v3.0 许可证 - 详见 [LICENSE](LICENSE) 文件

## 作者

OLC 
