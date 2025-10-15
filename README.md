# Fast URDF

一个用于快速处理和优化URDF（Unified Robot Description Format）文件的Python工具，提供图形化用户界面，支持URDF数据替换和网格简化功能。

## 功能特性

- 🤖 **URDF文件处理**：自动查找和处理URDF文件
- 📊 **Excel数据管理**：生成和管理连杆属性Excel文件
- 🔧 **数据替换**：基于Excel数据自动替换URDF文件中的连杆属性
- 🎨 **网格简化**：自动下载和简化STL网格文件以减小文件大小
- 🖥️ **图形化界面**：基于PyQt5的用户友好界面
- 🌐 **自动化下载**：自动下载Microsoft Edge WebDriver
- 📈 **实时进度反馈**：显示处理进度和日志信息

## 系统要求

- Python 3.6+
- Windows操作系统（支持Microsoft Edge）
- 至少2GB可用内存

## 安装说明

### 方法1：使用预构建的可执行文件

1. 下载release中的可执行文件
2. 双击运行 `Fast URDF.exe`

### 方法2：从源代码运行

#### 1. 克隆仓库

```bash
git clone https://github.com/yourusername/Fast-URDF.git
cd Fast-URDF
```

#### 2. 创建虚拟环境（推荐）

```bash
python -m venv myenv
myenv\Scripts\activate  # Windows
```

#### 3. 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 启动程序

```bash
python "Fast URDF.py"
```

或者运行批处理文件：

```bash
build.bat
```

### 工作流程

1. **选择工作目录**：选择包含URDF文件的工作目录
2. **设置参数**：
   - 文件大小阈值（MB）：设置需要简化的STL文件大小阈值
   - Edge驱动下载路径：指定WebDriver下载位置
   - Excel文件路径：指定连杆属性Excel文件位置

3. **选择功能**：
   - **仅URDF数据替换**：只替换URDF文件中的连杆属性
   - **URDF数据替换+网格简化**：同时进行数据替换和网格简化

### 目录结构要求

您的工作目录应该包含以下结构：

```text
工作目录/
├── urdf/              # 存放URDF文件
│   └── robot.urdf
├── meshes/            # 存放STL网格文件
│   ├── link1.stl
│   └── link2.stl
└── excel/             # 存放Excel属性文件（自动生成）
    └── links_properties.xlsx
```

## 依赖项

- `openpyxl` - Excel文件操作
- `requests` - HTTP请求和下载
- `selenium` - Web自动化（用于网格简化）
- `PyQt5` - 图形用户界面
- `pandas` - 数据处理

## 构建可执行文件

使用PyInstaller构建独立的可执行文件：

```bash
pip install pyinstaller
pyinstaller "Fast URDF.spec"
```

构建完成后，可执行文件将位于 `build/Fast URDF/` 目录中。

## 功能说明

### 1. Excel文件生成

- 自动扫描URDF文件中的连杆信息
- 生成包含连杆属性的Excel文件
- 支持质量、惯性、几何等属性的管理

### 2. URDF数据替换

- 读取Excel文件中的连杆属性数据
- 自动替换URDF文件中对应的属性值
- 支持质量、惯性矩阵、几何参数等的批量替换

### 3. 网格简化

- 自动检测超过阈值大小的STL文件
- 使用Web服务进行网格简化
- 自动下载和配置Microsoft Edge WebDriver

## 注意事项

- 请确保工作目录具有读写权限
- 网格简化功能需要网络连接
- 首次使用时会自动下载Edge WebDriver
- 建议在处理前备份原始URDF和STL文件

## 故障排除

### 常见问题

1. **找不到URDF文件**：
   - 确保URDF文件位于 `工作目录/urdf/` 文件夹中
   - 检查文件扩展名为 `.urdf`

2. **Excel文件错误**：
   - 检查Excel文件是否损坏
   - 确保文件格式正确

3. **网格简化失败**：
   - 检查网络连接
   - 确保Edge浏览器已安装
   - 尝试更新WebDriver

## 许可证

本项目采用MIT许可证，详见LICENSE文件。

## 贡献

欢迎提交Issue和Pull Request来改进这个项目！

## 更新日志

### v1.0.0

- 初始版本发布
- 支持URDF数据替换
- 支持网格简化
- 图形化用户界面
