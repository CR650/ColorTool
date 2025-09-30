// UGC工作表条件读取函数 - Custom_Fragile_Active_Color, Custom_Jump_Color, Custom_Jump_Active_Color
// 这些函数将被插入到themeManager.js中

/**
 * ========== Custom_Fragile_Active_Color 条件读取函数 ==========
 */

/**
 * 从源数据Custom_Fragile_Active_Color工作表读取字段值
 */
function findCustomFragileActiveColorValueFromSourceCustomFragileActiveColor(fieldName) {
    console.log(`=== 从源数据Custom_Fragile_Active_Color工作表查找字段: ${fieldName} ===`);

    if (!sourceData || !sourceData.workbook) {
        console.log('源数据不可用');
        return null;
    }

    try {
        const customFragileActiveColorSheet = sourceData.workbook.Sheets['Custom_Fragile_Active_Color'];
        if (!customFragileActiveColorSheet) {
            console.log('源数据中未找到Custom_Fragile_Active_Color工作表');
            return null;
        }

        const customFragileActiveColorData = XLSX.utils.sheet_to_json(customFragileActiveColorSheet, {
            header: 1,
            defval: '',
            raw: false
        });

        if (!customFragileActiveColorData || customFragileActiveColorData.length < 2) {
            console.log('源数据Custom_Fragile_Active_Color工作表数据不足');
            return null;
        }

        const headerRow = customFragileActiveColorData[0];
        const dataRow = customFragileActiveColorData[1];

        const fieldColumnIndex = headerRow.findIndex(col => col === fieldName);
        if (fieldColumnIndex === -1) {
            console.log(`源数据Custom_Fragile_Active_Color工作表中未找到字段: ${fieldName}`);
            return null;
        }

        const fieldValue = dataRow[fieldColumnIndex];
        if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
            console.log(`✅ 从源数据Custom_Fragile_Active_Color工作表找到: ${fieldName} = ${fieldValue}`);
            return fieldValue.toString();
        } else {
            console.log(`源数据Custom_Fragile_Active_Color工作表字段 ${fieldName} 值为空`);
            return null;
        }

    } catch (error) {
        console.error(`从源数据Custom_Fragile_Active_Color工作表读取字段 ${fieldName} 时出错:`, error);
        return null;
    }
}

/**
 * 从UGCTheme Custom_Fragile_Active_Color工作表读取字段值
 */
function findCustomFragileActiveColorValueFromUGCThemeCustomFragileActiveColor(fieldName, themeName) {
    console.log(`=== 从UGCTheme Custom_Fragile_Active_Color工作表查找字段: ${fieldName} (主题: ${themeName}) ===`);

    if (!ugcAllSheetsData || !ugcAllSheetsData['Custom_Fragile_Active_Color']) {
        console.log('UGCTheme Custom_Fragile_Active_Color数据未加载');
        return null;
    }

    try {
        const customFragileActiveColorData = ugcAllSheetsData['Custom_Fragile_Active_Color'];
        if (!customFragileActiveColorData || customFragileActiveColorData.length < 2) {
            console.log('UGCTheme Custom_Fragile_Active_Color工作表数据不足');
            return null;
        }

        const headerRow = customFragileActiveColorData[0];
        const fieldColumnIndex = headerRow.findIndex(col => col === fieldName);
        if (fieldColumnIndex === -1) {
            console.log(`UGCTheme Custom_Fragile_Active_Color工作表中未找到字段: ${fieldName}`);
            return null;
        }

        const notesColumnIndex = headerRow.findIndex(col => col === 'notes');
        if (notesColumnIndex === -1) {
            console.log('UGCTheme Custom_Fragile_Active_Color工作表中未找到notes列');
            return null;
        }

        for (let i = 1; i < customFragileActiveColorData.length; i++) {
            const row = customFragileActiveColorData[i];
            if (row[notesColumnIndex] === themeName) {
                const fieldValue = row[fieldColumnIndex];
                if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                    console.log(`✅ 从UGCTheme Custom_Fragile_Active_Color工作表找到: ${fieldName} = ${fieldValue} (主题: ${themeName})`);
                    return fieldValue.toString();
                } else {
                    console.log(`UGCTheme Custom_Fragile_Active_Color工作表中主题 ${themeName} 的字段 ${fieldName} 值为空`);
                    return null;
                }
            }
        }

        console.log(`UGCTheme Custom_Fragile_Active_Color工作表中未找到主题: ${themeName}`);
        return null;

    } catch (error) {
        console.error(`从UGCTheme Custom_Fragile_Active_Color工作表读取字段 ${fieldName} 时出错:`, error);
        return null;
    }
}

/**
 * Custom_Fragile_Active_Color字段条件读取逻辑
 */
function findCustomFragileActiveColorValueDirect(fieldName, isNewTheme = false, themeName = '') {
    console.log(`=== 直接映射查找Custom_Fragile_Active_Color字段值: ${fieldName} (新主题: ${isNewTheme}, 主题名: ${themeName}) ===`);

    if (!sourceData || !sourceData.workbook) {
        console.warn('源数据不可用');
        return null;
    }

    const statusInfo = parseStatusSheet(sourceData);
    console.log('Status状态信息:', statusInfo);

    if (!statusInfo.hasCustomFragileActiveColorField) {
        console.log('Status工作表中没有Custom_Fragile_Active_Color字段');
        if (!isNewTheme && themeName) {
            console.log('更新现有主题，从UGCTheme Custom_Fragile_Active_Color工作表读取');
            return findCustomFragileActiveColorValueFromUGCThemeCustomFragileActiveColor(fieldName, themeName);
        } else {
            console.log('新建主题且无Custom_Fragile_Active_Color字段，返回null');
            return null;
        }
    }

    const customFragileActiveColorStatus = statusInfo.customFragileActiveColorStatus;
    console.log(`Custom_Fragile_Active_Color状态: ${customFragileActiveColorStatus}`);

    if (customFragileActiveColorStatus === 1) {
        console.log('Custom_Fragile_Active_Color状态为1（有效），优先从源数据读取');
        const sourceValue = findCustomFragileActiveColorValueFromSourceCustomFragileActiveColor(fieldName);
        if (sourceValue !== null) {
            return sourceValue;
        }

        console.log('源数据中未找到，回退到UGCTheme Custom_Fragile_Active_Color工作表');
        if (themeName) {
            return findCustomFragileActiveColorValueFromUGCThemeCustomFragileActiveColor(fieldName, themeName);
        }
        return null;

    } else {
        console.log('Custom_Fragile_Active_Color状态为0（无效），忽略源数据，仅从UGCTheme读取');
        if (themeName) {
            return findCustomFragileActiveColorValueFromUGCThemeCustomFragileActiveColor(fieldName, themeName);
        }
        return null;
    }
}

/**
 * ========== Custom_Jump_Color 条件读取函数 ==========
 */

/**
 * 从源数据Custom_Jump_Color工作表读取字段值
 */
function findCustomJumpColorValueFromSourceCustomJumpColor(fieldName) {
    console.log(`=== 从源数据Custom_Jump_Color工作表查找字段: ${fieldName} ===`);

    if (!sourceData || !sourceData.workbook) {
        console.log('源数据不可用');
        return null;
    }

    try {
        const customJumpColorSheet = sourceData.workbook.Sheets['Custom_Jump_Color'];
        if (!customJumpColorSheet) {
            console.log('源数据中未找到Custom_Jump_Color工作表');
            return null;
        }

        const customJumpColorData = XLSX.utils.sheet_to_json(customJumpColorSheet, {
            header: 1,
            defval: '',
            raw: false
        });

        if (!customJumpColorData || customJumpColorData.length < 2) {
            console.log('源数据Custom_Jump_Color工作表数据不足');
            return null;
        }

        const headerRow = customJumpColorData[0];
        const dataRow = customJumpColorData[1];

        const fieldColumnIndex = headerRow.findIndex(col => col === fieldName);
        if (fieldColumnIndex === -1) {
            console.log(`源数据Custom_Jump_Color工作表中未找到字段: ${fieldName}`);
            return null;
        }

        const fieldValue = dataRow[fieldColumnIndex];
        if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
            console.log(`✅ 从源数据Custom_Jump_Color工作表找到: ${fieldName} = ${fieldValue}`);
            return fieldValue.toString();
        } else {
            console.log(`源数据Custom_Jump_Color工作表字段 ${fieldName} 值为空`);
            return null;
        }

    } catch (error) {
        console.error(`从源数据Custom_Jump_Color工作表读取字段 ${fieldName} 时出错:`, error);
        return null;
    }
}

/**
 * 从UGCTheme Custom_Jump_Color工作表读取字段值
 */
function findCustomJumpColorValueFromUGCThemeCustomJumpColor(fieldName, themeName) {
    console.log(`=== 从UGCTheme Custom_Jump_Color工作表查找字段: ${fieldName} (主题: ${themeName}) ===`);

    if (!ugcAllSheetsData || !ugcAllSheetsData['Custom_Jump_Color']) {
        console.log('UGCTheme Custom_Jump_Color数据未加载');
        return null;
    }

    try {
        const customJumpColorData = ugcAllSheetsData['Custom_Jump_Color'];
        if (!customJumpColorData || customJumpColorData.length < 2) {
            console.log('UGCTheme Custom_Jump_Color工作表数据不足');
            return null;
        }

        const headerRow = customJumpColorData[0];
        const fieldColumnIndex = headerRow.findIndex(col => col === fieldName);
        if (fieldColumnIndex === -1) {
            console.log(`UGCTheme Custom_Jump_Color工作表中未找到字段: ${fieldName}`);
            return null;
        }

        const notesColumnIndex = headerRow.findIndex(col => col === 'notes');
        if (notesColumnIndex === -1) {
            console.log('UGCTheme Custom_Jump_Color工作表中未找到notes列');
            return null;
        }

        for (let i = 1; i < customJumpColorData.length; i++) {
            const row = customJumpColorData[i];
            if (row[notesColumnIndex] === themeName) {
                const fieldValue = row[fieldColumnIndex];
                if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                    console.log(`✅ 从UGCTheme Custom_Jump_Color工作表找到: ${fieldName} = ${fieldValue} (主题: ${themeName})`);
                    return fieldValue.toString();
                } else {
                    console.log(`UGCTheme Custom_Jump_Color工作表中主题 ${themeName} 的字段 ${fieldName} 值为空`);
                    return null;
                }
            }
        }

        console.log(`UGCTheme Custom_Jump_Color工作表中未找到主题: ${themeName}`);
        return null;

    } catch (error) {
        console.error(`从UGCTheme Custom_Jump_Color工作表读取字段 ${fieldName} 时出错:`, error);
        return null;
    }
}

/**
 * Custom_Jump_Color字段条件读取逻辑
 */
function findCustomJumpColorValueDirect(fieldName, isNewTheme = false, themeName = '') {
    console.log(`=== 直接映射查找Custom_Jump_Color字段值: ${fieldName} (新主题: ${isNewTheme}, 主题名: ${themeName}) ===`);

    if (!sourceData || !sourceData.workbook) {
        console.warn('源数据不可用');
        return null;
    }

    const statusInfo = parseStatusSheet(sourceData);
    console.log('Status状态信息:', statusInfo);

    if (!statusInfo.hasCustomJumpColorField) {
        console.log('Status工作表中没有Custom_Jump_Color字段');
        if (!isNewTheme && themeName) {
            console.log('更新现有主题，从UGCTheme Custom_Jump_Color工作表读取');
            return findCustomJumpColorValueFromUGCThemeCustomJumpColor(fieldName, themeName);
        } else {
            console.log('新建主题且无Custom_Jump_Color字段，返回null');
            return null;
        }
    }

    const customJumpColorStatus = statusInfo.customJumpColorStatus;
    console.log(`Custom_Jump_Color状态: ${customJumpColorStatus}`);

    if (customJumpColorStatus === 1) {
        console.log('Custom_Jump_Color状态为1（有效），优先从源数据读取');
        const sourceValue = findCustomJumpColorValueFromSourceCustomJumpColor(fieldName);
        if (sourceValue !== null) {
            return sourceValue;
        }

        console.log('源数据中未找到，回退到UGCTheme Custom_Jump_Color工作表');
        if (themeName) {
            return findCustomJumpColorValueFromUGCThemeCustomJumpColor(fieldName, themeName);
        }
        return null;

    } else {
        console.log('Custom_Jump_Color状态为0（无效），忽略源数据，仅从UGCTheme读取');
        if (themeName) {
            return findCustomJumpColorValueFromUGCThemeCustomJumpColor(fieldName, themeName);
        }
        return null;
    }
}

/**
 * ========== Custom_Jump_Active_Color 条件读取函数 ==========
 */

/**
 * 从源数据Custom_Jump_Active_Color工作表读取字段值
 */
function findCustomJumpActiveColorValueFromSourceCustomJumpActiveColor(fieldName) {
    console.log(`=== 从源数据Custom_Jump_Active_Color工作表查找字段: ${fieldName} ===`);

    if (!sourceData || !sourceData.workbook) {
        console.log('源数据不可用');
        return null;
    }

    try {
        const customJumpActiveColorSheet = sourceData.workbook.Sheets['Custom_Jump_Active_Color'];
        if (!customJumpActiveColorSheet) {
            console.log('源数据中未找到Custom_Jump_Active_Color工作表');
            return null;
        }

        const customJumpActiveColorData = XLSX.utils.sheet_to_json(customJumpActiveColorSheet, {
            header: 1,
            defval: '',
            raw: false
        });

        if (!customJumpActiveColorData || customJumpActiveColorData.length < 2) {
            console.log('源数据Custom_Jump_Active_Color工作表数据不足');
            return null;
        }

        const headerRow = customJumpActiveColorData[0];
        const dataRow = customJumpActiveColorData[1];

        const fieldColumnIndex = headerRow.findIndex(col => col === fieldName);
        if (fieldColumnIndex === -1) {
            console.log(`源数据Custom_Jump_Active_Color工作表中未找到字段: ${fieldName}`);
            return null;
        }

        const fieldValue = dataRow[fieldColumnIndex];
        if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
            console.log(`✅ 从源数据Custom_Jump_Active_Color工作表找到: ${fieldName} = ${fieldValue}`);
            return fieldValue.toString();
        } else {
            console.log(`源数据Custom_Jump_Active_Color工作表字段 ${fieldName} 值为空`);
            return null;
        }

    } catch (error) {
        console.error(`从源数据Custom_Jump_Active_Color工作表读取字段 ${fieldName} 时出错:`, error);
        return null;
    }
}

/**
 * 从UGCTheme Custom_Jump_Active_Color工作表读取字段值
 */
function findCustomJumpActiveColorValueFromUGCThemeCustomJumpActiveColor(fieldName, themeName) {
    console.log(`=== 从UGCTheme Custom_Jump_Active_Color工作表查找字段: ${fieldName} (主题: ${themeName}) ===`);

    if (!ugcAllSheetsData || !ugcAllSheetsData['Custom_Jump_Active_Color']) {
        console.log('UGCTheme Custom_Jump_Active_Color数据未加载');
        return null;
    }

    try {
        const customJumpActiveColorData = ugcAllSheetsData['Custom_Jump_Active_Color'];
        if (!customJumpActiveColorData || customJumpActiveColorData.length < 2) {
            console.log('UGCTheme Custom_Jump_Active_Color工作表数据不足');
            return null;
        }

        const headerRow = customJumpActiveColorData[0];
        const fieldColumnIndex = headerRow.findIndex(col => col === fieldName);
        if (fieldColumnIndex === -1) {
            console.log(`UGCTheme Custom_Jump_Active_Color工作表中未找到字段: ${fieldName}`);
            return null;
        }

        const notesColumnIndex = headerRow.findIndex(col => col === 'notes');
        if (notesColumnIndex === -1) {
            console.log('UGCTheme Custom_Jump_Active_Color工作表中未找到notes列');
            return null;
        }

        for (let i = 1; i < customJumpActiveColorData.length; i++) {
            const row = customJumpActiveColorData[i];
            if (row[notesColumnIndex] === themeName) {
                const fieldValue = row[fieldColumnIndex];
                if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                    console.log(`✅ 从UGCTheme Custom_Jump_Active_Color工作表找到: ${fieldName} = ${fieldValue} (主题: ${themeName})`);
                    return fieldValue.toString();
                } else {
                    console.log(`UGCTheme Custom_Jump_Active_Color工作表中主题 ${themeName} 的字段 ${fieldName} 值为空`);
                    return null;
                }
            }
        }

        console.log(`UGCTheme Custom_Jump_Active_Color工作表中未找到主题: ${themeName}`);
        return null;

    } catch (error) {
        console.error(`从UGCTheme Custom_Jump_Active_Color工作表读取字段 ${fieldName} 时出错:`, error);
        return null;
    }
}

/**
 * Custom_Jump_Active_Color字段条件读取逻辑
 */
function findCustomJumpActiveColorValueDirect(fieldName, isNewTheme = false, themeName = '') {
    console.log(`=== 直接映射查找Custom_Jump_Active_Color字段值: ${fieldName} (新主题: ${isNewTheme}, 主题名: ${themeName}) ===`);

    if (!sourceData || !sourceData.workbook) {
        console.warn('源数据不可用');
        return null;
    }

    const statusInfo = parseStatusSheet(sourceData);
    console.log('Status状态信息:', statusInfo);

    if (!statusInfo.hasCustomJumpActiveColorField) {
        console.log('Status工作表中没有Custom_Jump_Active_Color字段');
        if (!isNewTheme && themeName) {
            console.log('更新现有主题，从UGCTheme Custom_Jump_Active_Color工作表读取');
            return findCustomJumpActiveColorValueFromUGCThemeCustomJumpActiveColor(fieldName, themeName);
        } else {
            console.log('新建主题且无Custom_Jump_Active_Color字段，返回null');
            return null;
        }
    }

    const customJumpActiveColorStatus = statusInfo.customJumpActiveColorStatus;
    console.log(`Custom_Jump_Active_Color状态: ${customJumpActiveColorStatus}`);

    if (customJumpActiveColorStatus === 1) {
        console.log('Custom_Jump_Active_Color状态为1（有效），优先从源数据读取');
        const sourceValue = findCustomJumpActiveColorValueFromSourceCustomJumpActiveColor(fieldName);
        if (sourceValue !== null) {
            return sourceValue;
        }

        console.log('源数据中未找到，回退到UGCTheme Custom_Jump_Active_Color工作表');
        if (themeName) {
            return findCustomJumpActiveColorValueFromUGCThemeCustomJumpActiveColor(fieldName, themeName);
        }
        return null;

    } else {
        console.log('Custom_Jump_Active_Color状态为0（无效），忽略源数据，仅从UGCTheme读取');
        if (themeName) {
            return findCustomJumpActiveColorValueFromUGCThemeCustomJumpActiveColor(fieldName, themeName);
        }
        return null;
    }
}

