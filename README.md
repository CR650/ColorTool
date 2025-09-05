# 颜色主题管理工具 (ColorTool Connect)

一个专为Unity项目设计的颜色主题配置管理工具，支持Excel文件的颜色数据处理和主题文件更新。

## 🎯 项目概述

本工具旨在简化Unity项目中颜色主题的管理流程，通过Web界面实现颜色配置的导入、处理和更新，支持RSC_Theme.xls和UGCTheme.xls文件的自动化处理。

## ✨ 主要功能

### 核心功能
- **颜色数据处理**：解析Excel源数据文件中的完整配色表
- **主题文件管理**：支持RSC_Theme.xls和UGCTheme.xls文件的读取和更新
- **数据映射**：智能匹配和映射颜色数据到目标主题文件
- **文件保存**：支持直接覆盖原文件或下载更新后的文件

### 技术特性
- **File System Access API**：在支持的浏览器中实现直接文件操作
- **拖拽上传**：支持文件拖拽上传，提升用户体验
- **数据预览**：实时预览Excel文件内容和Sheet数据
- **响应式设计**：适配桌面和移动设备
- **错误处理**：完善的错误提示和异常处理机制

## 🚀 快速开始

### 在线使用
访问 [GitHub Pages 部署地址](https://your-username.github.io/ColorToolConnectRC/) 直接使用。

### 本地部署
1. 克隆项目到本地
```bash
git clone https://github.com/your-username/ColorToolConnectRC.git
cd ColorToolConnectRC
```

2. 使用本地服务器运行
```bash
# 使用Python
python -m http.server 8000

# 使用Node.js
npx serve .

# 或直接用浏览器打开 index.html
```

## 📖 使用指南

### 基本使用流程
1. **选择源数据文件**：上传包含完整配色表的Excel文件
2. **选择主题文件**：选择Unity项目中的RSC_Theme.xls和UGCTheme.xls文件
3. **输入主题名称**：指定要创建或更新的主题名称
4. **处理数据**：点击处理按钮生成更新后的主题文件
5. **保存文件**：直接覆盖原文件或下载更新后的文件

### 文件格式要求
- **源数据文件**：Excel格式(.xlsx, .xls)，包含"完整配色表"工作表
- **主题文件**：Unity项目中的RSC_Theme.xls和UGCTheme.xls文件
- **输出格式**：统一使用.xls格式以确保与Unity工具兼容

### 浏览器兼容性
- **推荐浏览器**：Chrome 86+, Edge 86+ (支持File System Access API)
- **兼容浏览器**：Firefox, Safari (使用传统下载方式)
- **移动设备**：支持移动端浏览器访问

## 🏗️ 项目结构

```
ColorToolConnectRC/
├── index.html              # 主页面
├── css/
│   └── styles.css          # 样式文件
├── js/
│   ├── main.js            # 主模块
│   ├── version.js         # 版本管理
│   ├── utils.js           # 工具函数
│   ├── fileHandler.js     # 文件处理
│   ├── dataParser.js      # 数据解析
│   ├── tableRenderer.js   # 表格渲染
│   ├── exportUtils.js     # 导出工具
│   ├── themeManager.js    # 主题管理
│   ├── fileSystemAccess.js # 文件系统访问
│   └── enhancedFileManager.js # 增强文件管理
├── electron_app/          # Electron应用版本
├── python_backend/        # Python后端版本
└── README.md              # 项目文档
```

## 🔧 技术栈

### 前端技术
- **HTML5**：语义化标记和现代Web标准
- **CSS3**：响应式设计和现代样式
- **JavaScript ES6+**：模块化开发和现代语法
- **SheetJS (XLSX)**：Excel文件处理库

### 浏览器API
- **File System Access API**：直接文件操作
- **File API**：文件读取和处理
- **Drag and Drop API**：拖拽文件上传
- **LocalStorage API**：本地数据存储

### 开发工具
- **模块化架构**：清晰的代码组织结构
- **错误处理**：完善的异常捕获和处理
- **调试支持**：详细的控制台日志输出

## 🐛 问题排查

### 常见问题
1. **文件格式错误**：确保使用.xls格式的主题文件
2. **权限问题**：在Chrome/Edge中允许文件访问权限
3. **数据不匹配**：检查源数据文件是否包含"完整配色表"工作表
4. **浏览器兼容**：使用推荐的浏览器版本

### 错误代码
- **FileFormatError**：文件格式不支持
- **PermissionError**：文件访问权限不足
- **DataMappingError**：数据映射失败
- **NetworkError**：网络连接问题

## 🤝 贡献指南

欢迎提交Issue和Pull Request来改进项目。

### 开发环境设置
1. Fork项目到你的GitHub账户
2. 创建功能分支：`git checkout -b feature/new-feature`
3. 提交更改：`git commit -am 'Add new feature'`
4. 推送分支：`git push origin feature/new-feature`
5. 创建Pull Request

### 代码规范
- 使用ES6+语法
- 遵循模块化开发原则
- 添加适当的注释和文档
- 确保代码通过测试

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 🎨 功能特色

### 智能数据处理
- **自动识别**：智能识别Excel文件中的配色表工作表
- **数据验证**：自动验证数据格式和完整性
- **错误恢复**：提供详细的错误信息和修复建议
- **批量处理**：支持多个主题的批量更新

### 用户体验优化
- **实时预览**：即时预览文件内容和处理结果
- **进度指示**：清晰的处理进度和状态提示
- **操作引导**：详细的使用步骤和操作指南
- **快捷操作**：支持键盘快捷键和右键菜单

### 数据安全
- **本地处理**：所有数据处理在浏览器本地完成
- **隐私保护**：不上传任何文件到服务器
- **备份机制**：自动创建文件备份
- **版本控制**：支持文件版本管理和回滚

## 🔍 详细功能说明

### 文件处理流程
1. **文件验证**：检查文件格式、大小和结构
2. **数据解析**：提取Excel文件中的颜色数据
3. **数据映射**：将源数据映射到目标主题结构
4. **数据更新**：更新主题文件中的颜色配置
5. **文件生成**：生成更新后的主题文件

### 支持的数据格式
- **颜色格式**：RGB、HEX、HSL等多种颜色格式
- **数据类型**：数值、文本、公式等Excel数据类型
- **工作表**：支持多工作表文件处理
- **编码格式**：UTF-8、GBK等多种字符编码

### 高级功能
- **数据导出**：支持JSON、CSV等格式导出
- **模板管理**：保存和复用颜色主题模板
- **批量操作**：一次性处理多个文件
- **自定义映射**：用户自定义数据映射规则

## 📊 性能指标

### 处理能力
- **文件大小**：支持最大10MB的Excel文件
- **数据量**：可处理包含数万行数据的文件
- **响应时间**：通常在3秒内完成数据处理
- **内存使用**：优化的内存管理，避免浏览器卡顿

### 兼容性测试
- ✅ Chrome 86+ (完整功能)
- ✅ Edge 86+ (完整功能)
- ✅ Firefox 90+ (基础功能)
- ✅ Safari 14+ (基础功能)
- ✅ 移动端浏览器 (响应式界面)

## 🛠️ 开发指南

### 本地开发
```bash
# 克隆项目
git clone https://github.com/your-username/ColorToolConnectRC.git
cd ColorToolConnectRC

# 启动开发服务器
python -m http.server 8000
# 或使用 Node.js
npx serve .

# 访问应用
open http://localhost:8000
```

### 代码结构说明
- **模块化设计**：每个功能模块独立开发和测试
- **事件驱动**：基于事件的模块间通信
- **错误边界**：完善的错误捕获和处理机制
- **性能优化**：懒加载和异步处理优化

### 调试技巧
- 打开浏览器开发者工具查看详细日志
- 使用`App.Debug = true`开启调试模式
- 检查控制台错误信息和网络请求
- 使用浏览器的文件访问权限设置

## 📞 联系方式

如有问题或建议，请通过以下方式联系：
- 提交 [GitHub Issue](https://github.com/your-username/ColorToolConnectRC/issues)
- 发送邮件至：your-email@example.com
- 项目讨论：[GitHub Discussions](https://github.com/your-username/ColorToolConnectRC/discussions)

---

## 📋 版本历史

### v1.1.0 (2025-01-09)
**主要更新：Excel格式兼容性修复**
- 🔧 修复Excel文件格式兼容性问题
- 📝 统一使用.xls格式以兼容Unity工具
- 🐛 解决Apache POI HSSF库读取错误
- ✨ 添加版本显示功能
- 🚀 优化文件保存和下载流程
- 📱 改进响应式界面设计
- 🔍 增强错误提示和调试信息

**技术改进：**
- 所有`XLSX.write()`调用统一使用`bookType: 'xls'`
- 更新MIME类型为`application/vnd.ms-excel`
- 修复文件扩展名和下载文件名
- 添加版本管理模块和更新通知

### v1.0.0 (2025-01-08)
**初始版本发布**
- 🎉 项目初始化和基础架构搭建
- 📊 实现颜色主题数据处理功能
- 📁 支持RSC_Theme和UGCTheme文件管理
- 🔗 添加File System Access API支持
- 👀 实现数据预览和Sheet选择功能
- 📱 添加响应式界面设计
- 🎯 支持拖拽文件上传
- 🔄 实现数据同步和映射功能
- 🛠️ 完善错误处理和用户反馈
- 📖 添加使用指南和帮助文档

**核心功能：**
- Excel文件解析和处理
- 颜色数据映射和更新
- 多种文件保存方式
- 浏览器兼容性检测
- 模块化代码架构
