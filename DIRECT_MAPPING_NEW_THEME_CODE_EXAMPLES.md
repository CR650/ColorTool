# 直接映射模式新建主题 - 完整代码示例

## 📋 Light 配置的完整加载流程

### 场景1：间接映射模式（JSON模式）- ✅ 正确

```javascript
// ========== 第3443-3445行：主题名称输入处理函数 ==========
} else {
    // 非直接映射模式或无源数据，使用最后一个主题的Light配置作为默认值
    resetLightConfigToDefaults();  // ✅ 直接调用重置函数
}

// ========== 第1520行：resetLightConfigToDefaults 函数 ==========
function resetLightConfigToDefaults() {
    const lightDefaults = getLastThemeLightConfig();  // ✅ 获取第一个主题数据
    
    Object.entries(lightDefaults).forEach(([fieldId, defaultValue]) => {
        const input = document.getElementById(fieldId);
        if (input) {
            input.value = defaultValue;  // ✅ 设置UI值
        }
    });
}

// ========== 第1152行：getLastThemeLightConfig 函数 ==========
function getLastThemeLightConfig() {
    if (!rscAllSheetsData || !rscAllSheetsData['Light']) {
        return { lightMax: '0', lightDark: '0', ... };  // 硬编码默认值
    }

    const lightData = rscAllSheetsData['Light'];
    const lightHeaderRow = lightData[0];
    
    // ✅ 检查数据是否足够（需要至少6行）
    if (lightHeaderRow.findIndex(col => col === 'notes') === -1 || lightData.length <= 5) {
        return { lightMax: '0', lightDark: '0', ... };  // 硬编码默认值
    }

    // ✅ 读取第一个主题（行索引5，第6行）
    const firstRowIndex = 5;
    const firstRow = lightData[firstRowIndex];

    // ✅ 构建配置对象
    const firstThemeConfig = {};
    const lightFieldMapping = {
        'Max': 'lightMax',
        'Dark': 'lightDark',
        'Min': 'lightMin',
        'SpecularLevel': 'lightSpecularLevel',
        'Gloss': 'lightGloss',
        'SpecularColor': 'lightSpecularColor'
    };

    Object.entries(lightFieldMapping).forEach(([columnName, fieldId]) => {
        const columnIndex = lightHeaderRow.findIndex(col => col === columnName);
        if (columnIndex !== -1) {
            const value = firstRow[columnIndex];
            firstThemeConfig[fieldId] = (value !== undefined && value !== null && value !== '') 
                ? value.toString() 
                : '0';
        }
    });

    return firstThemeConfig;  // ✅ 返回第一个主题的数据
}
```

**结果：** ✅ UI显示第一个主题（赛博1）的Light配置数据

---

### 场景2：直接映射模式 - ❌ 错误

```javascript
// ========== 第3439-3445行：主题名称输入处理函数 ==========
if (currentMappingMode === 'direct' && sourceData && sourceData.workbook) {
    console.log('🔧 新建主题（直接映射模式）：尝试从源数据加载Light配置到UI');
    loadExistingLightConfig(inputValue, true);  // ❌ 传递 isNewTheme=true
} else {
    resetLightConfigToDefaults();
}

// ========== 第11780行：loadExistingLightConfig 函数 ==========
function loadExistingLightConfig(themeName, isNewTheme = false) {
    console.log('加载Light配置:', themeName);
    console.log('是否新建主题:', isNewTheme);
    console.log('当前映射模式:', currentMappingMode);

    const isDirectMode = currentMappingMode === 'direct';

    if (isDirectMode) {
        console.log('直接映射模式：优先从源数据Light工作表读取配置显示');

        const lightFieldMapping = {
            'Max': 'lightMax',
            'Dark': 'lightDark',
            'Min': 'lightMin',
            'SpecularLevel': 'lightSpecularLevel',
            'Gloss': 'lightGloss',
            'SpecularColor': 'lightSpecularColor'
        };

        let hasSourceData = false;

        // ❌ 问题：这里调用 findLightValueDirect，但它在 isNewTheme=true 时总是返回 null
        Object.entries(lightFieldMapping).forEach(([lightColumn, inputId]) => {
            const directValue = findLightValueDirect(lightColumn, isNewTheme, themeName);
            // ❌ directValue 总是 null（当 isNewTheme=true 时）

            const input = document.getElementById(inputId);
            if (input) {
                if (directValue !== null && directValue !== undefined && directValue !== '') {
                    input.value = directValue;
                    hasSourceData = true;
                } else {
                    // ❌ 回退到默认值，但这个默认值是错的
                    const lightDefaults = getLastThemeLightConfig();
                    const defaultValue = lightDefaults[inputId] || '';
                    input.value = defaultValue;
                    console.log(`⚠️ 直接映射模式Light字段条件读取失败，使用默认值: ${lightColumn} = ${defaultValue}`);
                }
            }
        });

        if (hasSourceData) {
            console.log('✅ 直接映射模式：成功从源数据加载Light配置');
            return;
        } else {
            console.log('⚠️ 直接映射模式：未能从源数据获取Light配置，回退到RSC_Theme读取');
        }
    }
    // ... 回退逻辑
}

// ========== 第5846行：findLightValueDirect 函数（问题所在）==========
function findLightValueDirect(lightField, isNewTheme = false, themeName = '') {
    console.log(`=== 直接映射查找Light字段值: ${lightField} (新主题: ${isNewTheme}, 主题名: ${themeName}) ===`);

    if (!sourceData || !sourceData.workbook) {
        console.warn('源数据不可用');
        return null;  // ❌ 返回 null
    }

    const statusInfo = parseStatusSheet(sourceData);
    console.log('Status状态信息:', statusInfo);

    if (!statusInfo.hasLightField) {
        console.warn('Status工作表中没有Light字段，根据主题类型处理');

        if (!isNewTheme && themeName) {  // ❌ 当 isNewTheme=true 时，这个条件是 false
            // 更新现有主题模式：直接从RSC_Theme Light工作表读取
            const rscLightValue = findLightValueFromRSCThemeLight(lightField, themeName);
            if (rscLightValue) return rscLightValue;
        }

        // ❌ 新建主题时直接返回 null
        console.log(`⚠️ 无Light字段，${isNewTheme ? '新建主题' : '现有主题'}未找到Light字段值: ${lightField}`);
        return null;
    }

    const isLightValid = statusInfo.isLightValid;
    console.log(`Light字段状态: ${isLightValid ? '有效(1)' : '无效(0)'}`);

    if (isLightValid) {
        // Light状态为有效(1)
        console.log('Light状态有效，优先从源数据Light工作表查找');

        const sourceLightValue = findLightValueFromSourceLight(lightField);
        if (sourceLightValue) {
            console.log(`✅ 从源数据Light工作表找到: ${lightField} = ${sourceLightValue}`);
            return sourceLightValue;
        }

        if (!isNewTheme && themeName) {  // ❌ 当 isNewTheme=true 时，这个条件是 false
            // 更新现有主题模式：回退到RSC_Theme Light工作表
            const rscLightValue = findLightValueFromRSCThemeLight(lightField, themeName);
            if (rscLightValue) return rscLightValue;
        }

        // ❌ 新建主题时直接返回 null
        console.log(`⚠️ Light状态有效但未找到Light字段值: ${lightField}`);
        return null;

    } else {
        // Light状态为无效(0)
        console.log('Light状态无效，忽略源数据Light工作表');

        if (!isNewTheme && themeName) {  // ❌ 当 isNewTheme=true 时，这个条件是 false
            // 更新现有主题模式：直接从RSC_Theme Light工作表读取
            const rscLightValue = findLightValueFromRSCThemeLight(lightField, themeName);
            if (rscLightValue) return rscLightValue;
        }

        // ❌ 新建主题时直接返回 null
        console.log(`⚠️ Light状态无效，${isNewTheme ? '新建主题' : '现有主题'}未找到Light字段值: ${lightField}`);
        return null;
    }
}
```

**结果：** ❌ `findLightValueDirect` 总是返回 `null`，最终使用硬编码默认值

---

## 🔴 问题总结

**直接映射模式下新建主题时的问题：**

1. **调用链不同**：
   - 间接映射：`resetLightConfigToDefaults()` → `getLastThemeLightConfig()` ✅
   - 直接映射：`loadExistingLightConfig(themeName, true)` → `findLightValueDirect()` ❌

2. **条件读取函数的缺陷**：
   - `findLightValueDirect()` 中所有的数据读取逻辑都被 `if (!isNewTheme && themeName)` 保护
   - 当 `isNewTheme=true` 时，这个条件永远是 `false`
   - 所以新建主题时，条件读取逻辑完全不工作

3. **相同问题存在于其他三个函数**：
   - `findColorInfoValueDirect()` - ColorInfo配置
   - `findFloodLightValueDirect()` - FloodLight配置
   - `findVolumetricFogValueDirect()` - VolumetricFog配置

---

## ✅ 修复方案

需要修改四个条件读取函数，在 `isNewTheme=true` 时，也要读取第一个主题的数据：

```javascript
// 修改 findLightValueDirect 函数的逻辑
if (isNewTheme) {
    // 新建主题模式：读取第一个主题的数据
    if (isLightValid) {
        // Light状态有效：优先从源数据Light工作表读取第一个主题
        const sourceLightValue = findLightValueFromSourceLight(lightField);
        if (sourceLightValue) return sourceLightValue;
        
        // 回退到RSC_Theme Light工作表的第一个主题
        const rscLightValue = findLightValueFromRSCThemeLightFirstTheme(lightField);
        if (rscLightValue) return rscLightValue;
    } else {
        // Light状态无效：直接从RSC_Theme Light工作表读取第一个主题
        const rscLightValue = findLightValueFromRSCThemeLightFirstTheme(lightField);
        if (rscLightValue) return rscLightValue;
    }
    return null;
}
```

这样就能保证直接映射模式下新建主题时，也能读取到第一个主题的正确数据。

