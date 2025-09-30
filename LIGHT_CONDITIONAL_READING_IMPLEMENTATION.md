# Light工作表条件读取机制实施文档

## 📋 实施概述

本次实施在ColorTool Connect项目的直接映射模式下，成功实现了Light工作表的条件读取逻辑，与Color字段的处理方式保持一致。

## 🎯 核心修改内容

### 1. 扩展parseStatusSheet函数
**文件**: `js/themeManager.js` (行 2527-2609)
**修改内容**: 添加Light字段状态解析逻辑
```javascript
// 查找Light列
const lightColumnIndex = headers.findIndex(header => {
    if (!header) return false;
    const headerStr = header.toString().trim().toUpperCase();
    return headerStr === 'LIGHT';
});

let lightStatus = 0;
let hasLightField = false;
if (lightColumnIndex !== -1) {
    const lightStatusValue = statusRow[lightColumnIndex];
    lightStatus = parseInt(lightStatusValue) || 0;
    hasLightField = true;
    console.log(`Light字段状态: ${lightStatus} (原始值: "${lightStatusValue}")`);
}
```

### 2. 新增Light数据读取函数

#### findLightValueFromSourceLight函数
**位置**: `js/themeManager.js` (行 3838-3886)
**功能**: 从源数据Light工作表读取Light字段值
**支持字段**: Max, Dark, Min, SpecularLevel, Gloss, SpecularColor

#### findLightValueFromRSCThemeLight函数
**位置**: `js/themeManager.js` (行 3894-3963)
**功能**: 从RSC_Theme Light工作表读取Light字段值
**查找方式**: 通过notes列匹配主题名称

#### findLightValueDirect函数
**位置**: `js/themeManager.js` (行 3972-4048)
**功能**: 实现Light字段的条件读取逻辑
**优先级策略**:
- **更新现有主题**: 源数据Light工作表 → RSC_Theme Light工作表 → 默认值
- **新建主题**: 源数据Light工作表 → 默认值

### 3. 修改applyLightConfigToRow函数
**位置**: `js/themeManager.js` (行 4785-4857)
**修改内容**: 
- 添加themeName和isNewTheme参数
- 在直接映射模式下使用条件读取逻辑
- 在非直接映射模式下保持原有的用户配置读取方式

### 4. 更新函数调用链
**修改的函数**:
- `addNewRowToSheet`: 添加isNewTheme参数并传递给applyLightConfigToRow
- `updateExistingRowInSheet`: 传递themeName和isNewTheme=false给applyLightConfigToRow
- `processRSCAdditionalSheets`: 传递isNewTheme参数给addNewRowToSheet

### 5. 更新映射模式指示器
**位置**: `js/themeManager.js` (行 2669-2682)
**修改内容**: 在UI中显示Light状态信息
```javascript
const lightStatus = additionalInfo.lightStatus !== undefined ?
    (additionalInfo.lightStatus === 1 ? '有效' : '无效') : '未知';
mappingModeContent.innerHTML = `
    <div class="mapping-mode-title">
        <span class="mapping-mode-icon">🎯</span>直接映射模式
    </div>
    <div class="mapping-mode-description">
        检测到Status工作表，Color状态: ${colorStatus}，Light状态: ${lightStatus}，支持${additionalInfo.fieldCount || 0}个直接字段映射
    </div>
`;
```

### 6. 扩展公共接口
**位置**: `js/themeManager.js` (行 8706-8711)
**新增测试函数**:
- `findLightValueDirect`
- `findLightValueFromSourceLight`
- `findLightValueFromRSCThemeLight`

## 🔧 Light字段状态判断逻辑

### Status工作表中Light字段的处理
- **Light列存在且值为1**: Light表数据有效，启用Light表数据读取
- **Light列存在且值为0**: Light表数据无效，忽略Light表数据
- **Light列不存在**: 视为Light表数据无效

### 数据读取优先级

#### 更新现有主题模式（isNewTheme=false）
1. **Light状态=1（有效）**: 源数据Light工作表 → RSC_Theme Light工作表 → 默认值
2. **Light状态=0（无效）**: RSC_Theme Light工作表 → 默认值
3. **Status工作表无Light字段**: RSC_Theme Light工作表 → 默认值

#### 新建主题模式（isNewTheme=true）
1. **Light状态=1（有效）**: 源数据Light工作表 → 默认值
2. **Light状态=0（无效）**: 默认值
3. **Status工作表无Light字段**: 默认值

## 📊 支持的Light字段

| 字段名称 | 描述 | 默认值来源 |
|---------|------|-----------|
| Max | 最大光照值 | 上一个主题的lightMax |
| Dark | 暗部光照值 | 上一个主题的lightDark |
| Min | 最小光照值 | 上一个主题的lightMin |
| SpecularLevel | 镜面反射级别 | 上一个主题的lightSpecularLevel |
| Gloss | 光泽度 | 上一个主题的lightGloss |
| SpecularColor | 镜面反射颜色 | 上一个主题的lightSpecularColor |

## 🧪 测试验证

### 测试文件
- `test-light-conditional-reading.html`: 完整的Light条件读取测试页面

### 测试场景
1. **Status工作表 + Light状态=1**: 验证从源数据Light工作表读取
2. **Status工作表 + Light状态=0**: 验证忽略源数据Light工作表
3. **Status工作表无Light字段**: 验证根据主题类型的回退逻辑
4. **新建主题 vs 更新现有主题**: 验证不同主题类型的处理差异

## ✅ 实施保证

### 兼容性保证
- ✅ **JSON间接映射模式完全不受影响**
- ✅ **向后兼容性得到保证**
- ✅ **其他字段（Color、FloodLight、VolumetricFog等）处理逻辑不受影响**

### 功能保证
- ✅ **Light字段的处理逻辑与Color字段保持一致**
- ✅ **错误处理和日志输出机制保持完整**
- ✅ **UI界面正常显示Light状态信息**
- ✅ **支持多种Light字段类型的条件读取**

### 性能保证
- ✅ **仅在直接映射模式下启用条件读取逻辑**
- ✅ **非直接映射模式保持原有性能**
- ✅ **合理的错误处理避免性能问题**

## 🔄 与Color字段的一致性

Light字段的条件读取逻辑完全遵循Color字段的设计模式：
- 相同的Status工作表解析逻辑
- 相同的优先级策略
- 相同的主题类型处理方式
- 相同的错误处理和日志格式
- 相同的公共接口暴露方式

## 📝 使用说明

### 开发者使用
```javascript
// 直接调用Light条件读取函数
const lightValue = App.ThemeManager.findLightValueDirect('Max', false, 'MyTheme');

// 从源数据Light工作表读取
const sourceValue = App.ThemeManager.findLightValueFromSourceLight('SpecularLevel');

// 从RSC_Theme Light工作表读取
const rscValue = App.ThemeManager.findLightValueFromRSCThemeLight('Gloss', 'MyTheme');
```

### 用户使用
1. 准备包含Status工作表的源数据文件
2. 在Status工作表中设置Light字段状态（0或1）
3. 如果Light状态为1，确保源数据包含Light工作表
4. 系统将自动根据Light状态进行条件读取

## 🎉 实施完成

Light工作表的条件读取机制已成功实施，与Color字段的处理方式保持完全一致，为ColorTool Connect项目的直接映射模式提供了更加灵活和强大的数据处理能力。
