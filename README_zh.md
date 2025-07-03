# Excel-PPT映射器 (ExcelPPTSliderMapper)

一个强大的Python工具，用于将Excel数据自动映射到PowerPoint模板中，支持PTML（PPT Template Markup Language）标记语言。

## 🌟 功能特性

### 核心功能
- **自动数据映射**: 从Excel文件读取数据，自动替换PPT模板中的占位符
- **多种内容支持**: 支持文本、表格、图表、图片等多种PowerPoint内容类型
- **PTML标记语言**: 使用简洁的 `${变量名}` 语法标记需要替换的内容
- **智能文件处理**: 自动检测和处理文件占用问题
- **进程管理**: 自动管理PowerPoint进程，避免冲突

### 支持的数据类型
- **文本数据**: 普通文本、百分比、数值等
- **日期数据**: 自动格式化日期显示
- **表格数据**: 动态表格内容替换
- **图表数据**: 支持柱状图、折线图、组合图表等
- **图片数据**: 动态图片替换

### 高级特性
- **不区分大小写**: 支持键名的大小写不敏感匹配
- **百分比智能处理**: 自动识别和格式化百分比数据
- **组合图表支持**: 支持柱状图+折线图的组合图表
- **文件安全保存**: 带重试机制的安全文件保存
- **详细日志输出**: 完整的处理过程日志

## 📋 系统要求

### Python环境
- Python 3.6+
- Windows操作系统（需要COM组件支持）

### 依赖库
```
win32com.client  # Windows COM组件
pandas          # 数据处理
openpyxl        # Excel文件读取
psutil          # 进程管理
```

### Microsoft Office
- Microsoft PowerPoint 2016+ 
- Microsoft Excel 2016+

## 🚀 安装说明

1. **克隆项目**
```bash
git clone https://github.com/your-repo/ExcelPPTSliderMapper.git
cd ExcelPPTSliderMapper
```

2. **安装依赖**
```bash
pip install -r requirements.txt
```

3. **安装pywin32**
```bash
pip install pywin32
```

## 📖 使用方法

### 基本用法

```python
from src.excel_ppt_mapper.excel_ppt_mapper import process_ptml_template, read_excel_template

# 文件路径
ppt_template = "template.pptx"
excel_data = "data.xlsx" 
output_file = "result.pptx"

# 读取Excel数据
template_data = read_excel_template(excel_data)

# 处理PPT模板
success = process_ptml_template(
    ppt_path=ppt_template,
    template_data=template_data,
    output_path=output_file,
    page_numbers=[1, 2, 3]  # 可选：指定处理的页面
)

if success:
    print("✅ PPT处理完成")
else:
    print("❌ PPT处理失败")
```

### PTML标记语法

在PowerPoint模板中使用以下标记：

```
${变量名}  # 直接替换，不进行任何转换
```

**示例：**
- `${公司名称}` → "无仓科技有限公司"
- `${增长率}` → "25%"
- `${报告日期}` → "2024-01-15"

## 📊 Excel数据格式

### 文本数据 (Text工作表)
| Key | Value |
|-----|-------|
| 公司名称 | 无仓科技有限公司 |
| 增长率 | 25% |
| 营收 | 1000万元 |

### 日期数据 (Dates工作表)
| Key | Value |
|-----|-------|
| 报告日期 | 2024-01-15 |
| 截止日期 | 2024-12-31 |

### 表格数据 (Tables工作表)
| Column | Description |
|--------|-------------|
| table_name | 表格名称 |
| header | 表头行 |
| data | 数据行 |

### 图表数据 (combo_charts工作表)
| 图表名称 | 分类 | 系列类型 | 系列名称 | 数值 | 图表类型 | 标题 |
|---------|-----|---------|---------|-----|---------|-----|
| 销售图表 | Q1 | column | 销售额 | 100 | combo | 季度销售分析 |
| 销售图表 | Q1 | line | 增长率 | 15 | combo | 季度销售分析 |

### 图片数据 (Images工作表)
| Key | Value |
|-----|-------|
| 公司Logo | ./images/logo.png |
| 产品图片 | ./images/product.jpg |

## 🔧 API文档

### 主要函数

#### `process_ptml_template(ppt_path, template_data, output_path, page_numbers)`
处理PPT模板的主函数

**参数：**
- `ppt_path` (str): PPT模板文件路径
- `template_data` (Dict): 模板数据字典
- `output_path` (str, 可选): 输出文件路径
- `page_numbers` (List[int], 可选): 要处理的页面编号列表

**返回值：**
- `bool`: 处理是否成功

#### `read_excel_template(excel_path)`
从Excel文件读取模板数据

**参数：**
- `excel_path` (str): Excel文件路径

**返回值：**
- `Dict[str, Any]`: 模板数据字典

### 辅助函数

#### `check_file_in_use(file_path)`
检查文件是否被占用

#### `close_powerpoint_processes()`
关闭所有PowerPoint进程

#### `safe_save_presentation(pres, output_path, max_retries)`
安全保存演示文稿，带重试机制

## 💡 使用示例

### 示例1：处理完整PPT
```python
# 处理所有页面
template_data = read_excel_template("data.xlsx")
process_ptml_template("template.pptx", template_data, "output.pptx")
```

### 示例2：处理指定页面
```python
# 只处理第1-3页
template_data = read_excel_template("data.xlsx")
process_ptml_template(
    "template.pptx", 
    template_data, 
    "output.pptx", 
    page_numbers=[1, 2, 3]
)
```

### 示例3：文本替换
**PPT模板中：**
```
公司：${公司名称}
增长率：${增长率}
```

**Excel数据：**
| Key | Value |
|-----|-------|
| 公司名称 | 无仓科技 |
| 增长率 | 25% |

**结果：**
```
公司：无仓科技
增长率：25%
```

## ⚠️ 注意事项

### 文件操作
1. **文件权限**: 确保有足够的权限读写文件
2. **文件占用**: 程序会自动处理文件占用问题
3. **备份重要文件**: 建议在处理前备份原始文件

### PowerPoint相关
1. **版本兼容性**: 建议使用Office 2016及以上版本
2. **宏安全性**: 可能需要调整PowerPoint的宏安全设置
3. **进程清理**: 程序会自动清理PowerPoint进程

### 数据格式
1. **编码问题**: Excel文件建议使用UTF-8编码
2. **数据类型**: 注意数值、文本、日期的格式一致性
3. **特殊字符**: 避免在键名中使用特殊字符

### 性能优化
1. **大文件处理**: 大型PPT文件处理可能需要较长时间
2. **内存使用**: 复杂图表可能占用较多内存
3. **并发限制**: 不建议同时运行多个实例

## 🐛 故障排除

### 常见问题

**Q: 提示"文件正在使用中"**
A: 程序会自动关闭PowerPoint进程，如果仍有问题，请手动关闭所有Office程序

**Q: 标记没有被替换**
A: 检查Excel中的键名是否与PPT中的标记完全匹配（支持大小写不敏感）

**Q: 图表数据更新失败**
A: 确保Excel中的图表数据格式正确，特别是数值类型

**Q: 程序运行缓慢**
A: 可能是由于PPT文件较大或图表较复杂，请耐心等待

### 调试模式
程序提供详细的日志输出，可以通过查看控制台输出来诊断问题。

## 📄 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

## 🤝 贡献

欢迎提交Issues和Pull Requests来帮助改进这个项目！

## 📞 支持

如果您遇到任何问题或有功能建议，请：
1. 查看本文档的故障排除部分
2. 搜索现有的Issues
3. 创建新的Issue

---

**注意**: 本工具需要在Windows环境下运行，并依赖Microsoft Office套件。
