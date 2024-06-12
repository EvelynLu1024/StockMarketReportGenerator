# 某险资权益投资部晨会纪要生成器

本项目从wind提取市场数据，生成A股、港股、美股在某确定区间的市场表现。。以下是项目的详细信息和使用说明。

## 目录

- [安装](#dump_az)
- [使用说明](#dump_sysm)
- [项目结构](#dump_xmjg)
- [代码解释](#dump_dmjs)
- [贡献](#dump_gx)
- [许可证](#dump_xkz)

## <span id=dump_az>安装</span>

在使用本项目之前，请确保已安装所需的Python库。您可以使用以下命令安装`python-docx`库：

```bash
pip install python-docx
```

## <span id=dump_sysm>使用说明</span>

- 运行main.py文件，生成word文件。

- 目前只支持wind接口。

- 通过mt_weekly.yaml文件更改基础信息。

## <span id=dump_xmjg>项目结构</span>

```plaintext
MTWeeklyCode/
├── main.py                # 主文件，用于执行生成Word文档
├── morningtalk_weekly.py  # 主要代码文件，包含类和主要逻辑
├── mt_weekly.yaml         # 配置文件
├── README.md              # 项目说明文件
└── draft.py               # 不包含类的草稿文件
```
## <span id=dump_dmjs>代码解释</span>

## MorningTalkWeekly 类代码介绍

### 导入库

```python
import os
import warnings
import pandas as pd
from datetime import datetime
import sys
from WindPy import w
import yaml
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn
```
这些导入语句引入了必要的库，包括用于文件操作、日期处理、数据处理和文档生成的库。

### MorningTalkWeekly 类
## 类结构

```plaintext
MorningTalkWeekly
├── __init__(self, start_date, end_date, yaml_path, output_dir)
├── load_yaml(self)
├── get_zdfweekly_w(self, code)
├── sign_transformation(value)
├── describe_indus(self, df)
├── describe_wind_index(self, df, top_n=15)
├── generate_word_report(self)
└── get_paragraphs(self)
```
#### 函数说明
- \_\_init\_\_(self, start_date, end_date, yaml_path, output_dir): 初始化类的实例变量，并启动WindPy数据服务。

```python
class MorningTalkWeekly:
    def __init__(self, start_date, end_date, yaml_path, output_dir):
        start_date  # 开始日期（包括）
        end_date  # 结束日期（包括）
        yaml_path  # mt_weekly.yaml文件路径
        output_dir  # 输出文件的路径
```

- load_yaml(self): 从指定路径加载YAML配置文件并返回其内容。
- get_zdfweekly_w(self, code): 从WindPy获取指定代码的周涨跌幅数据并返回处理后的DataFrame。
- sign_transformation(value): 根据涨跌幅值返回相应的描述字符串。
- describe_indus(self, df): 根据涨跌幅排序并描述前5个上涨和下跌的行业。
- describe_wind_index(self, df, top_n=15): 根据涨跌幅排序并描述前15个上涨和下跌的Wind概念。
- generate_word_report(self): 生成一个包含标题、日期和市场概况的Word文档，并将其保存到指定目录。
- get_paragraphs(self): 生成并返回包含A股、港股和美股市场描述的段落列表。



### 示例:
在主程序中实例化 MorningTalkWeekly 类并生成报告：

```bash
if __name__ == "__main__":
    morningtalk_weekly = MorningTalkWeekly(
        start_date="20240603",
        end_date="20240607",
        yaml_path='mt_weekly.yaml',
        output_dir='.'
    )
    morningtalk_weekly.generate_word_report()
```

## <span id=dump_gx> 贡献</span>
欢迎任何形式的贡献！以下是一些可以帮助你开始贡献的方法：

#### 如何贡献

1. **报告问题**: 如果你发现了任何错误或者有任何建议，请通过[EvelynLu1024@outlook.com](EvelynLu1024@outlook.com)联系。
2. **提交请求**: 如果你已经解决了一个问题或者添加了一个新功能，请提交一个Pull Request。
3. **改进文档**: 如果你发现文档有需要改进的地方，请随时提出或者直接修改并提交。

#### 代码规范

- 确保代码风格与项目保持一致。
- 提交前请运行所有测试并确保它们都通过。

#### 联系作者

- 如果你有任何问题或需要帮助，请联系 [EvelynLu1024@outlook.com](EvelynLu1024@outlook.com)。

感谢你的贡献！


## <span id=dump_xkz> 许可证</span>
MIT License

Copyright (c) 2024 YingLu
