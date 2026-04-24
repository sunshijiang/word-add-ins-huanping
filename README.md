# Word 表格格式调整工具

一个用于快速调整 Word 文档中表格格式的 Office 加载项。

## 功能特性

### 🎨 预设样式
- **优雅风格** - 浅灰底纹，细边框，宋体
- **商务风格** - 蓝色主题，微软雅黑，粗边框
- **学术风格** - 仿宋，小五号，规范格式
- **极简风格** - 无样式，仅基础边框
- **彩色风格** - 橙色主题

### 🔧 手动调整
- 字体设置（名称、字号）
- 段落对齐（左/中/右/两端）
- 行距和段距调整
- 边框宽度与颜色自定义
- 表格样式选择

### ✨ 智能检测
- 自动检测光标是否在表格内
- 实时状态提示
- 条件显示调整面板

## 安装部署

### 开发环境
```bash
# 安装依赖
npm install

# 启动开发服务器
npm run dev-server

# 启动调试
npm start
```

### 生产构建
```bash
# 构建生产版本
npm run build

# 构建开发版本
npm run build:dev

# 监听文件变化
npm run watch
```

### 代码质量
```bash
# 代码检查
npm run lint

# 自动修复
npm run lint:fix

# 格式化代码
npm run prettier
```

## 项目结构

```
word-add-ins-huanping/
├── src/
│   ├── taskpane/          # 任务窗格
│   │   ├── taskpane.html  # 主界面
│   │   ├── taskpane.js    # 核心逻辑
│   │   └── taskpane.css   # 样式
│   └── commands/          # 命令功能
│       ├── commands.html
│       └── commands.js
├── assets/                # 图标资源
├── manifest.xml           # 加载项清单
├── webpack.config.js      # 构建配置
├── package.json           # 项目配置
└── README.md              # 说明文档
```

## 技术栈

- **Office.js** - Office 加载项 API
- **Webpack 5** - 模块打包
- **Babel** - JavaScript 转译
- **HTML/CSS/JS** - 原生前端技术

## 使用说明

1. 在 Word 中将光标放入表格任意单元格
2. 在"开始"选项卡找到"表格格式调整工具"按钮
3. 点击按钮打开任务窗格
4. 选择预设样式或手动调整参数
5. 点击"应用"按钮完成格式调整

## 开发指南

### 添加新预设
在 `src/taskpane/taskpane.js` 的 `PRESETS` 对象中添加：

```javascript
custom: {
    style: "TableStyleLight1",
    font: { name: "宋体", size: 10.5, color: "black" },
    alignment: "center",
    lineSpacing: 15,
    shading: { color: "F2F2F2" }
}
```

### 修改界面
编辑 `src/taskpane/taskpane.html` 和 `taskpane.css`

## 配置说明

### manifest.xml
- 应用 ID: `0c0a2bda-4e36-40e9-b4ae-cb3175420603`
- 支持主机：Word 文档
- 权限：读写文档

### webpack.config.js
- 开发端口：3000
- HTTPS 自动证书
- 热重载支持

## 许可证

MIT License

## 版本信息

- 版本：0.0.1
- 支持：Word Desktop (IE 11+)
- 最后更新：2026-04-24
