# v1.7.6 发布说明

**发布日期**: 2025-10-29  
**版本号**: 1.7.6  
**构建号**: 20251029001  
**类型**: Patch（补丁版本）

---

## 📋 概述

本版本修复了**直接映射模式下新建主题时的UI初始值加载问题**。在直接映射模式中，新建主题时各个UI配置面板（Light、ColorInfo、FloodLight、VolumetricFog等）显示的默认值不正确。

### 问题现象
- ✅ **间接映射模式（JSON模式）**：新建主题时UI显示正确（读取第一个主题的数据）
- ❌ **直接映射模式**：新建主题时UI显示错误（显示硬编码默认值而非第一个主题的数据）

---

## 🔧 核心修复

### 根本原因

所有条件读取函数中的数据读取逻辑都被以下条件保护：

```javascript
if (!isNewTheme && themeName) {
    // 读取数据的逻辑
}
```

当 `isNewTheme=true` 时，这个条件永远是 `false`，导致：
- ❌ 所有数据读取逻辑都被跳过
- ❌ 函数直接返回 `null`
- ❌ 最终使用硬编码默认值

### 修复方案

为所有条件读取函数添加对 `isNewTheme=true` 的支持：

```javascript
if (isNewTheme) {
    // ✅ 新建主题模式：从RSC_Theme/UGCTheme读取第一个主题的数据
    const value = findValueFromFirstTheme(field);
    if (value) return value;
} else if (themeName) {
    // 更新现有主题模式：按原有逻辑处理
    const value = findValueFromExistingTheme(field, themeName);
    if (value) return value;
}
```

---

## 📝 修复的函数

### RSC_Theme 配置（4个函数 + 4个辅助函数）

1. **Light 光照配置**
   - 修复函数：`findLightValueDirect()`
   - 新增辅助函数：`findLightValueFromRSCThemeLightFirstTheme()`

2. **ColorInfo 钻石颜色配置**
   - 修复函数：`findColorInfoValueDirect()`
   - 新增辅助函数：`findColorInfoValueFromRSCThemeColorInfoFirstTheme()`
   - 特殊规则：钻石颜色字段始终读取第一个主题

3. **FloodLight 泛光灯配置**
   - 修复函数：`findFloodLightValueDirect()`
   - 新增辅助函数：`findFloodLightValueFromRSCThemeFloodLightFirstTheme()`

4. **VolumetricFog 体积雾配置**
   - 修复函数：`findVolumetricFogValueDirect()`
   - 新增辅助函数：`findVolumetricFogValueFromRSCThemeVolumetricFogFirstTheme()`

### UGCTheme 配置（5个函数）

1. `findCustomGroundColorValueDirect()` - 地面颜色
2. `findCustomFragileColorValueDirect()` - 易碎块颜色
3. `findCustomFragileActiveColorValueDirect()` - 易碎块激活颜色
4. `findCustomJumpColorValueDirect()` - 跳跃块颜色
5. `findCustomJumpActiveColorValueDirect()` - 跳跃块激活颜色

---

## ✨ 修复效果

| 配置类型 | 修复前 | 修复后 |
|---------|--------|--------|
| **Light** | 新建主题返回null | 读取第一个主题 ✅ |
| **ColorInfo** | 新建主题返回null | 读取第一个主题 ✅ |
| **FloodLight** | 新建主题返回null | 读取第一个主题 ✅ |
| **VolumetricFog** | 新建主题返回null | 读取第一个主题 ✅ |
| **UGC配置** | 新建主题返回null | 按Status状态处理 ✅ |

---

## 🧪 测试验证

### 测试场景：直接映射模式 + 新建主题

1. 加载源数据文件（包含Color工作表）
2. 选择Unity项目文件夹
3. 输入新主题名称（如"测试主题1"）
4. 查看UI各配置面板的初始值
5. **预期结果**：UI显示第一个主题（赛博1）的数据 ✅

### 验证清单

- [ ] Light配置面板显示第一个主题的光照数据
- [ ] ColorInfo配置面板显示第一个主题的钻石颜色数据
- [ ] FloodLight配置面板显示第一个主题的泛光灯数据
- [ ] VolumetricFog配置面板显示第一个主题的体积雾数据
- [ ] UGC配置面板按Status状态正确处理

---

## 📊 影响范围

### 仅影响
- ✅ 直接映射模式下新建主题时的UI初始值显示

### 不影响
- ✅ 间接映射模式（JSON模式）
- ✅ 更新现有主题的逻辑
- ✅ 文件保存和数据处理流程
- ✅ 其他工作表的处理逻辑

---

## 🛡️ 向后兼容性

- ✅ 完全兼容现有的映射模式检测机制
- ✅ 完全兼容现有的Status状态检查机制
- ✅ 不影响其他工作表的处理逻辑
- ✅ 不影响新建主题的其他流程
- ✅ 不影响更新现有主题的流程

---

## 📚 相关文档

- README.md - 已更新版本历史和功能说明
- js/version.js - 已更新版本信息
- js/themeManager.js - 已修复所有条件读取函数

---

## 🎯 下一步

建议用户在直接映射模式下测试新建主题功能，验证UI初始值是否正确显示第一个主题的数据。

