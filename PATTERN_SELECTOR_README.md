# 🎨 Floor Pattern Selector - 图案选择器

## 📋 项目概述

Floor Pattern Selector 是一个交互式图案选择工具，专门用于从 `Floor_Pattern.png` 图片中选择单个图案。该工具提供了直观的用户界面，支持点击选择、快速定位、图案预览和导出功能。

## ✨ 功能特性

### 🎯 核心功能
- **交互式选择**: 点击图片上的任意区域选择对应图案
- **网格显示**: 可切换显示/隐藏网格线，便于精确定位
- **实时预览**: 选中图案的实时预览显示
- **详细信息**: 显示图案的位置、坐标、尺寸等详细信息

### 🛠️ 操作功能
- **快速选择**: 通过行列数字快速定位图案
- **图案导出**: 将选中的图案导出为独立的PNG文件
- **信息复制**: 复制图案信息到剪贴板
- **重置选择**: 一键清除当前选择

### 🎨 用户体验
- **响应式设计**: 适配不同屏幕尺寸
- **视觉反馈**: 选中状态的高亮显示
- **状态提示**: 实时状态更新和操作反馈
- **模态框确认**: 重要操作的确认对话框

## 📁 文件结构

```
├── pattern-selector.html          # 完整的图案选择器页面
├── pattern-integration-example.html  # 集成示例页面
├── js/
│   ├── patternSelector.js         # 核心图案选择逻辑
│   └── pattern-selector-ui.js     # UI控制器
├── css/
│   └── pattern-selector.css       # 样式文件
└── Texture/
    └── Floor_Pattern.png          # 图案源文件
```

## 🚀 快速开始

### 1. 基础使用

直接打开 `pattern-selector.html` 即可使用完整功能：

```bash
# 在浏览器中打开
open pattern-selector.html
```

### 2. 集成到现有项目

#### 引入必要文件

```html
<!-- CSS样式 -->
<link rel="stylesheet" href="css/pattern-selector.css">

<!-- JavaScript模块 -->
<script src="js/patternSelector.js"></script>
<script src="js/pattern-selector-ui.js"></script>
```

#### HTML结构

```html
<div id="patternSelectorContainer" class="pattern-container">
    <canvas id="patternCanvas" class="pattern-canvas"></canvas>
</div>
<div id="patternPreview" class="pattern-preview"></div>
<div id="patternInfo" class="pattern-info"></div>
```

#### JavaScript初始化

```javascript
// 初始化图案选择器
document.addEventListener('DOMContentLoaded', function() {
    if (window.App && window.App.PatternSelector) {
        window.App.PatternSelector.init();
    }
    
    if (window.App && window.App.PatternSelectorUI) {
        window.App.PatternSelectorUI.init();
    }
});
```

## 🔧 API 参考

### PatternSelector 模块

#### 初始化
```javascript
App.PatternSelector.init()
```

#### 选择图案
```javascript
// 通过图案对象选择
App.PatternSelector.selectPattern(pattern)

// 通过ID选择
App.PatternSelector.selectPatternById('pattern_2_3')

// 通过行列选择 (0基索引)
App.PatternSelector.selectPatternByPosition(2, 3)
```

#### 获取信息
```javascript
// 获取当前选中的图案
const selected = App.PatternSelector.getSelectedPattern()

// 获取所有图案网格
const grid = App.PatternSelector.getPatternGrid()

// 获取配置信息
const config = App.PatternSelector.getPatternConfig()
```

#### 重置选择
```javascript
App.PatternSelector.resetSelection()
```

### PatternSelectorUI 模块

#### 显示通知
```javascript
App.PatternSelectorUI.showNotification('消息内容', 'success', 3000)
// 类型: 'success', 'error', 'warning', 'info'
```

#### 更新状态
```javascript
App.PatternSelectorUI.updateStatus('新状态文本')
App.PatternSelectorUI.updateSelectedCount(1)
```

#### 按钮控制
```javascript
App.PatternSelectorUI.enableActionButtons()
App.PatternSelectorUI.disableActionButtons()
```

## 📡 事件系统

### 监听图案选择事件

```javascript
document.addEventListener('patternSelected', function(event) {
    const pattern = event.detail.pattern;
    console.log('选中图案:', pattern);
    
    // 处理选择逻辑
    handlePatternSelection(pattern);
});
```

### 监听图案使用事件

```javascript
document.addEventListener('patternUsed', function(event) {
    const pattern = event.detail.pattern;
    console.log('使用图案:', pattern);
    
    // 处理使用逻辑
    applyPattern(pattern);
});
```

### 监听网格切换事件

```javascript
document.addEventListener('toggleGrid', function(event) {
    const showGrid = event.detail.showGrid;
    console.log('网格显示状态:', showGrid);
});
```

## ⚙️ 配置选项

### 图案网格配置

```javascript
const PATTERN_CONFIG = {
    imageWidth: 1024,      // 图片宽度
    imageHeight: 1024,     // 图片高度
    gridCols: 8,           // 网格列数
    gridRows: 8,           // 网格行数
    patternWidth: 128,     // 单个图案宽度
    patternHeight: 128     // 单个图案高度
};
```

### 自定义样式

可以通过CSS变量自定义主题色彩：

```css
:root {
    --primary-color: #007bff;
    --secondary-color: #6c757d;
    --success-color: #28a745;
    --danger-color: #dc3545;
    --warning-color: #ffc107;
    --info-color: #17a2b8;
}
```

## 🎯 使用场景

### 1. 游戏开发
- 地板纹理选择
- UI图标选择
- 装饰元素选择

### 2. 设计工具
- 图案库管理
- 素材选择器
- 模板选择

### 3. 内容管理
- 资源管理系统
- 图片分类工具
- 素材预览

## 🔍 示例代码

### 完整集成示例

参考 `pattern-integration-example.html` 文件，该文件展示了：

- 简化版的图案选择器实现
- 基本的点击选择功能
- 图案预览和信息显示
- 导出和使用功能

### 自定义处理函数

```javascript
function handlePatternSelection(pattern) {
    // 自定义选择处理逻辑
    console.log(`选择了图案: ${pattern.name}`);
    console.log(`位置: 行${pattern.row + 1}, 列${pattern.col + 1}`);
    
    // 可以在这里添加自己的业务逻辑
    // 例如：更新游戏场景、保存用户选择等
}

function applyPattern(pattern) {
    // 自定义应用逻辑
    console.log(`应用图案: ${pattern.name}`);
    
    // 例如：将图案应用到游戏对象
    // gameObject.setTexture(pattern);
}
```

## 🐛 故障排除

### 常见问题

1. **图片加载失败**
   - 检查 `Floor_Pattern.png` 文件路径
   - 确保图片文件存在且可访问

2. **点击无响应**
   - 检查JavaScript控制台是否有错误
   - 确保所有必要的文件都已正确引入

3. **样式显示异常**
   - 检查CSS文件是否正确引入
   - 确保没有样式冲突

4. **功能不完整**
   - 确保按正确顺序初始化模块
   - 检查DOM元素是否存在

### 调试技巧

```javascript
// 检查模块是否正确加载
console.log('PatternSelector:', window.App.PatternSelector);
console.log('PatternSelectorUI:', window.App.PatternSelectorUI);

// 检查初始化状态
console.log('PatternSelector ready:', App.PatternSelector.isReady());
console.log('PatternSelectorUI ready:', App.PatternSelectorUI.isReady());
```

## 📝 更新日志

### v1.0.0 (2025-09-08)
- ✨ 初始版本发布
- 🎯 基础图案选择功能
- 🎨 完整的UI界面
- 📱 响应式设计支持
- 🔧 模块化架构设计

## 📄 许可证

本项目采用 MIT 许可证，详情请参阅 LICENSE 文件。

## 🤝 贡献

欢迎提交 Issue 和 Pull Request 来改进这个项目！

---

**享受使用图案选择器！** 🎨✨
