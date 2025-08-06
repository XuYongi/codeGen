# Excel数据可视化工具

这是一个用于可视化Excel数据的前端项目，特别处理了第六列中的换行符，并支持各个维度的筛选过滤。

## 项目特点

- 读取并解析Excel文件（支持 .xlsx 和 .xls 格式）
- 特别处理第六列中的换行符，统一为 `\n` 格式
- 支持按各个维度进行筛选过滤
- 支持多种数据展示模式：
  - 单条数据竖向展示
  - JSON格式展示（支持折叠展开）
  - 表格模式展示
- 支持逐条浏览数据（分页功能）
- 响应式设计，适配不同屏幕尺寸
- 特别处理 `extra_params` 字段，提取其中的 `prompt` 字段单独展示

## 项目结构

```
project/
│
├── index.html          # 主页面
├── styles/
│   └── main.css        # 样式文件
├── scripts/
│   ├── main.js         # 主脚本文件
│   ├── excelProcessor.js # Excel数据处理模块
│   └── dataVisualizer.js # 数据可视化模块
├── package.json        # 项目配置文件
└── README.md           # 项目说明文件
```

## 功能说明

### Excel数据处理
- 使用 `xlsx` 库读取和解析Excel文件
- 特别处理第六列中的换行符，统一为 `\n` 格式
- 将数据转换为易于处理的JSON格式

### 数据筛选
- 支持按各个维度（列）进行筛选
- 实时筛选，输入即生效
- 支持部分匹配（包含匹配）

### 数据展示模式

#### 表格模式
- 只显示关键字段：
  - `prompt`（从 `ide_codemate_model_request.extra_params` 中提取）
  - `model_output`
  - `ide_codemate_model_feedback.accept_content`
  - `accept_type`

#### 单条数据模式
- 竖向展示关键字段
- `prompt` 字段从 `extra_params` 中提取并单独显示
- 提供"显示其他字段"按钮，可展开查看完整数据
- `extra_params` 在"其他字段"中以JSON格式展示

#### JSON模式
- 以JSON格式展示完整数据
- 支持折叠/展开功能
- 格式化显示，易于阅读

### 分页功能
- 支持逐条浏览数据
- 上一条/下一条导航
- 显示当前条目位置

## 安装和运行

1. 确保你已经安装了 [Node.js](https://nodejs.org/)

2. 安装项目依赖：
   ```
   npm install
   ```

3. 启动开发服务器：
   ```
   npm start
   ```

   这将启动一个本地开发服务器，并在浏览器中打开项目。

## 使用方法

1. 点击"选择Excel文件"按钮上传你的Excel文件
2. 数据加载完成后，会显示筛选控件、视图控制和第一条数据
3. 使用以下方式操作：
   - 在筛选控件中选择或输入关键字进行筛选
   - 使用"上一条"/"下一条"按钮浏览数据
   - 切换不同的视图模式查看数据
   - 在JSON模式下，可以折叠/展开JSON内容
   - 在单条数据模式下，点击"显示其他字段"查看更多数据

## 技术栈

- HTML5
- CSS3
- JavaScript (ES6 Modules)
- [SheetJS (xlsx)](https://github.com/SheetJS/sheetjs) - 用于处理Excel文件

## 自定义

你可以根据需要修改以下文件：

- `index.html` - 修改页面结构和内容
- `styles/main.css` - 修改页面样式
- `scripts/main.js` - 主要的交互逻辑
- `scripts/excelProcessor.js` - Excel数据处理逻辑
- `scripts/dataVisualizer.js` - 数据可视化逻辑

## 许可证

本项目采用 MIT 许可证。