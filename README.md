# 打印助手 (Print Assistant)

一个用于生成76*130mm面单的Python工具，支持从Excel文件读取数据并生成PDF格式的打印标签。

## 功能特点

- 📊 支持Excel文件导入（.xlsx格式）
- 📄 生成76*130mm标准面单格式
- 🖨️ 支持直接打印或保存为PDF
- 🎨 自动格式化日期为yyyy-mm-dd
- 📱 跨平台支持（Windows/Mac）
- 🔤 支持中文字体显示

## 安装依赖

```bash
pip3 install pandas reportlab openpyxl
```

## 字体文件

请下载 [NotoSansSC-VariableFont_wght.ttf](https://fonts.google.com/noto/specimen/Noto+Sans+SC) 字体文件，并放置在脚本同目录下。

## 使用方法

1. 运行脚本：
```bash
python3 print_label.py
```

2. 选择Excel文件（支持.xlsx格式）

3. 选择PDF保存路径

4. 自动生成面单PDF文件

## Excel文件格式要求

Excel文件应包含以下列：

| 列名 | 说明 | 示例 |
|------|------|------|
| 订单日期 | 订单日期 | 2025-05-13 |
| 包装件数 | 件数 | 1 |
| 装载号 | 装载号 | 925051200805-2 |
| 买家姓名 | 收货人姓名（含电话） | 张良祥 15863339004 |
| 地址 | 收货地址 | 山东省日照市东港区... |
| 品名 | 商品名称 | DS8575B-A~1820x2020mm~床+淡奶白+高脚 |
| 顾客名字 | 顾客姓名 | 范克娟日照市五莲县 |

## 面单格式说明

- **尺寸**：76mm × 130mm
- **边距**：四周各1mm
- **仓库名**：固定为"河北顾家家居仓"
- **内容布局**：
  - 标题行：仓库名称（合并4列）
  - 日期行：日期、件数
  - 装载号行：装载号（合并后3列）
  - 收货信息行：收货人、地址（地址合并后2列）
  - 品名行：品名（合并后3列）
  - 顾客行：顾客姓名（合并后3列）

## 技术栈

- Python 3.x
- pandas：Excel文件读取
- reportlab：PDF生成
- openpyxl：Excel文件解析
- tkinter：文件选择对话框

## 系统要求

- Python 3.6+
- Windows 10+ 或 macOS 10.12+
- 至少50MB可用磁盘空间

## 故障排除

### 常见问题

1. **字体显示为方块**
   - 确保 `NotoSansSC-VariableFont_wght.ttf` 文件在脚本目录下
   - 检查字体文件是否损坏

2. **Excel读取失败**
   - 确保Excel文件为.xlsx格式
   - 检查文件是否被其他程序占用

3. **PDF生成失败**
   - 检查输出路径是否有写入权限
   - 确保磁盘空间充足

### 错误代码

- `NameError: name 'SimpleDocTemplate' is not defined`
  - 重新安装reportlab：`pip3 install --upgrade reportlab`

- `TTFError: postscript outlines are not supported`
  - 使用TTF格式字体，避免OTF格式

## 更新日志

- v1.0.0：初始版本，支持基础面单生成
- v1.1.0：优化表格布局，支持中文显示
- v1.2.0：调整列宽和行高，避免分页问题

## 许可证

本项目采用MIT许可证。

## 联系方式

如有问题或建议，请通过以下方式联系：
- 提交Issue
- 发送邮件

---

**注意**：请确保Excel文件格式正确，字体文件完整，以获得最佳使用体验。
