### Openpyxl
#### 1.import
```py
import openpyxl as vb
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
```

#### 2.创建工作簿
```py
wb = Workbook()
wb.save(path)
```

#### 3.打开工作簿
```py
wb = load_workbook('test.xlsx', data_only=True)
```

#### 4.工作簿的属性和方法
| 属性或方法     | 作用                                    |
|----------------|-----------------------------------------|
| active         | 获取当前活跃的工作表                    |
| worksheets     | 以列表形式返回所有的工作表（可迭代对象) |
| sheetnames     | 获取工作簿中的工作表（列表）            |
| wb['sheet1']   | 通过工作表名称获取工作表对象            |
| creat_sheet    | 创建一个空的工作表                      |
| remove         | 删除一个工作表对象                      |
| copy_worksheet | 复制一个工作表                          |

#### 5.工作表的属性和方法
| 属性和方法                                        | 作用                                                               |
|---------------------------------------------------|--------------------------------------------------------------------|
| title                                             | 给工作表命名                                                       |
| insert_cols(idx=2, amount=5)                      | 从第2列开始插入5列                                                 |
| insert_rows(idx=2, amount=5)                      | 从第2行开始插入5行                                                 |
| delete_rows(2, 5)                                 | 从第2行开始删除5行                                                 |
| delete_cols(2, 5)                                 | 从第2列开始删除5列                                                 |
| max_row                                           | 表格的最大行                                                       |
| max_column                                        | 表的的最大列                                                       |
| rows                                              | 按行获取单元格（Cell对象）生成器                                   |
| columns                                           | 按列获取单元格（Cell对象）生成器                                   |
| freeze_panes                                      | 冻结窗格                                                           |
| iter_rows                                         | 按行获取所有单元格，内置属性有min_row, max_row, min_col, max_col   |
| iter_columns                                      | 按列获取所有的单元格                                               |
| append                                            | 在工作表末尾添加数据                                               |
| merge_cells                                       | 合并单元格,属性有start_row, start_column, end_row, end_column      |
| unmerge_cells                                     | 移除合并的单元格                                                   |
| column_dimensions_group("A", "D", hidden=False)   | 按列进行分组                                                       |
| row_dimensions_group(1, 4, hidden=False           | 按行进行分组                                                       |
| row_dimensions[1].height                          | 设置行高                                                           |
| column_dimensions['B'].width                      | 设置列宽                                                           |

#### 6.字母转换
```py
# 数字转字母
get_column_letter(26)

# 字母转数字
column_index_from_string('D')
```

#### 7.批注
```py
ws['A1'].comment = vb.comment.Comment('test', 'se7en')
```

#### 8.Font
```py
from openpyxl.styles import Font

wb['A1'].font = Font(name=u'微软雅黑', bold=True, size=12)
```

| 属性       | 解释                                                                                     |
|------------|------------------------------------------------------------------------------------------|
| name       | 字体名称                                                                                 |
| size       | 字号大小                                                                                 |
| bold       | 字体粗细                                                                                 |
| italic     | 字体倾斜                                                                                 |
| vertAlign  | 上标和下标，默认为None                                                                   |
| underlinie | 默认None,single单下划线, double双下划线,singleAccounting会计用单下划线，doubleAccounting |
| strike     | 删除线                                                                                   |
| color      | 字体颜色                                                                                 |

#### 9.Alignment
| 属性          | 解释                                   |
|---------------|----------------------------------------|
| horizontal    | 水平对其方式，general常规，left, right |
| vertical      | 垂直对其方式, center, bottom           |
| text_rotation | 文本旋转角度                           |
| wrap_text     | 是否自动换行                           |
| indent        | 缩进                                   |


#### 10.Side and Border
| Side属性 | 解释                                                         |
|----------|--------------------------------------------------------------|
| style    | 连线样式，thin, double, hair, dashed, dashDot, thick, dotted |
| color    | 边框颜色                                                     |

| Border属性 | 解释 |
|------------|------|
| left       |      |
| right      |      |
| top        |      |
| bottom     |      |

```py
from openpyxl.styles import Side, Border

side = Side(style='thin', color='FF000000')
border = Border(left=side, right=side, top=side, bottom=side)
ws['A1'].border = border
```

#### 11.PatternFill
| 属性        | 解释                           |
|-------------|--------------------------------|
| patternType | None, solid                    |
| start_color | 'FF27E85B'                     |
| end_color   | 'FF27E85B'                     |
| fgColor     | getattr(colors, color.upper()) |
