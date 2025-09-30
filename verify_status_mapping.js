/**
 * 验证Status工作表条件映射机制的脚本
 * 用于测试修改后的直接映射模式逻辑
 */

// 模拟XLSX库的基本功能
const mockXLSX = {
    utils: {
        book_new: () => ({ SheetNames: [], Sheets: {} }),
        aoa_to_sheet: (data) => ({ data: data }),
        book_append_sheet: (wb, ws, name) => {
            wb.SheetNames.push(name);
            wb.Sheets[name] = ws;
        },
        sheet_to_json: (sheet, options) => {
            if (options && options.header === 1) {
                return sheet.data || [];
            }
            return [];
        }
    }
};

// 模拟全局XLSX对象
global.XLSX = mockXLSX;

// 创建测试用的工作簿
function createTestWorkbook(sheetData) {
    const wb = mockXLSX.utils.book_new();
    
    Object.keys(sheetData).forEach(sheetName => {
        const ws = mockXLSX.utils.aoa_to_sheet(sheetData[sheetName]);
        mockXLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    
    return wb;
}

// 模拟detectMappingMode函数（基于修改后的逻辑）
function detectMappingMode(sourceData) {
    console.log('=== 开始检测映射模式 ===');

    if (!sourceData || !sourceData.workbook) {
        console.log('源数据无效，使用JSON映射模式');
        return 'json';
    }

    const sheetNames = sourceData.workbook.SheetNames;
    console.log('源数据工作表列表:', sheetNames);

    // 第一优先级：检查是否包含"完整配色表"工作表
    const hasCompleteColorSheet = sheetNames.includes('完整配色表');
    if (hasCompleteColorSheet) {
        console.log('✅ 找到"完整配色表"工作表，使用JSON间接映射模式');
        return 'json';
    }

    // 第二优先级：检查是否包含"Status"工作表
    const hasStatusSheet = sheetNames.includes('Status');
    if (hasStatusSheet) {
        console.log('✅ 找到"Status"工作表，使用直接映射模式');

        // 验证Status工作表是否有有效数据
        try {
            const statusSheet = sourceData.workbook.Sheets['Status'];
            const statusData = mockXLSX.utils.sheet_to_json(statusSheet, {
                header: 1,
                defval: '',
                raw: false
            });

            if (!statusData || statusData.length < 2) {
                console.log('⚠️ Status工作表数据不足，回退到JSON映射模式');
                return 'json';
            }

            const headers = statusData[0];
            console.log('Status工作表表头:', headers);

            // 简化检测：只要有表头和数据就启用直接映射
            if (headers && headers.length > 0) {
                console.log(`✅ 检测到直接映射模式：Status工作表包含${headers.length}个字段`);
                return 'direct';
            } else {
                console.log('⚠️ Status工作表表头为空，回退到JSON映射模式');
                return 'json';
            }

        } catch (error) {
            console.error('读取Status工作表时出错:', error);
            console.log('⚠️ Status工作表读取失败，回退到JSON映射模式');
            return 'json';
        }
    }

    // 默认情况：没有找到特定工作表，使用JSON映射模式
    console.log('未找到"完整配色表"或"Status"工作表，使用JSON映射模式');
    return 'json';
}

// 模拟parseStatusSheet函数
function parseStatusSheet(sourceData) {
    console.log('=== 开始解析Status工作表 ===');

    if (!sourceData || !sourceData.workbook) {
        console.warn('源数据无效，无法解析Status工作表');
        return { colorStatus: 0, hasColorField: false, error: '源数据无效' };
    }

    try {
        const statusSheet = sourceData.workbook.Sheets['Status'];
        if (!statusSheet) {
            console.warn('Status工作表不存在');
            return { colorStatus: 0, hasColorField: false, error: 'Status工作表不存在' };
        }

        const statusData = mockXLSX.utils.sheet_to_json(statusSheet, {
            header: 1,
            defval: '',
            raw: false
        });

        if (!statusData || statusData.length < 2) {
            console.warn('Status工作表数据不足');
            return { colorStatus: 0, hasColorField: false, error: 'Status工作表数据不足' };
        }

        const headers = statusData[0];
        const statusRow = statusData[1]; // 第二行是状态行

        console.log('Status工作表表头:', headers);
        console.log('Status工作表状态行:', statusRow);

        // 查找Color列的索引
        const colorColumnIndex = headers.findIndex(header => {
            if (!header) return false;
            const headerStr = header.toString().trim().toUpperCase();
            return headerStr === 'COLOR';
        });

        if (colorColumnIndex === -1) {
            console.log('Status工作表中未找到Color列');
            return { colorStatus: 0, hasColorField: false, error: '未找到Color列' };
        }

        const colorStatusValue = statusRow[colorColumnIndex];
        const colorStatus = parseInt(colorStatusValue) || 0;

        console.log(`Color字段状态: ${colorStatus} (原始值: "${colorStatusValue}")`);

        const result = {
            colorStatus: colorStatus,
            hasColorField: true,
            colorColumnIndex: colorColumnIndex,
            headers: headers,
            statusRow: statusRow,
            isColorValid: colorStatus === 1
        };

        console.log('Status工作表解析结果:', result);
        return result;

    } catch (error) {
        console.error('解析Status工作表时出错:', error);
        return { colorStatus: 0, hasColorField: false, error: error.message };
    }
}

// 测试函数
function runTests() {
    console.log('🧪 开始Status工作表条件映射机制验证测试\n');

    // 测试1: Status工作表 + Color状态=1
    console.log('=== 测试1: Status工作表 + Color状态=1 ===');
    const test1Data = {
        'Status': [
            ['Color', 'Light', 'FloodLight'],  // 表头
            [1, 0, 1]                          // 状态行：Color=1(有效)
        ],
        'Color': [
            ['P1', 'P2', 'G1'],               // Color工作表表头
            ['FF0000', '00FF00', '0000FF']    // 颜色数据
        ]
    };
    
    const workbook1 = createTestWorkbook(test1Data);
    const sourceData1 = { workbook: workbook1, fileName: 'test1.xlsx' };
    
    const mode1 = detectMappingMode(sourceData1);
    const status1 = parseStatusSheet(sourceData1);
    
    console.log(`结果: 映射模式=${mode1}, Color状态=${status1.colorStatus}, 有效性=${status1.isColorValid}`);
    console.log(mode1 === 'direct' && status1.isColorValid ? '✅ 测试1通过' : '❌ 测试1失败');
    console.log('');

    // 测试2: Status工作表 + Color状态=0
    console.log('=== 测试2: Status工作表 + Color状态=0 ===');
    const test2Data = {
        'Status': [
            ['Color', 'Light', 'FloodLight'],
            [0, 1, 0]  // Color=0(无效)
        ],
        'Color': [
            ['P1', 'P2', 'G1'],
            ['FF0000', '00FF00', '0000FF']
        ]
    };
    
    const workbook2 = createTestWorkbook(test2Data);
    const sourceData2 = { workbook: workbook2, fileName: 'test2.xlsx' };
    
    const mode2 = detectMappingMode(sourceData2);
    const status2 = parseStatusSheet(sourceData2);
    
    console.log(`结果: 映射模式=${mode2}, Color状态=${status2.colorStatus}, 有效性=${status2.isColorValid}`);
    console.log(mode2 === 'direct' && !status2.isColorValid ? '✅ 测试2通过' : '❌ 测试2失败');
    console.log('');

    // 测试3: 完整配色表工作表（JSON映射模式）
    console.log('=== 测试3: 完整配色表工作表（JSON映射模式）===');
    const test3Data = {
        '完整配色表': [
            ['颜色代码', '16进制值'],
            ['P1', 'FF0000'],
            ['P2', '00FF00']
        ]
    };
    
    const workbook3 = createTestWorkbook(test3Data);
    const sourceData3 = { workbook: workbook3, fileName: 'test3.xlsx' };
    
    const mode3 = detectMappingMode(sourceData3);
    
    console.log(`结果: 映射模式=${mode3}`);
    console.log(mode3 === 'json' ? '✅ 测试3通过' : '❌ 测试3失败');
    console.log('');

    // 测试4: 无特殊工作表（默认JSON映射）
    console.log('=== 测试4: 无特殊工作表（默认JSON映射）===');
    const test4Data = {
        'Sheet1': [
            ['普通数据'],
            ['测试']
        ]
    };
    
    const workbook4 = createTestWorkbook(test4Data);
    const sourceData4 = { workbook: workbook4, fileName: 'test4.xlsx' };
    
    const mode4 = detectMappingMode(sourceData4);
    
    console.log(`结果: 映射模式=${mode4}`);
    console.log(mode4 === 'json' ? '✅ 测试4通过' : '❌ 测试4失败');
    console.log('');

    console.log('🎯 所有测试完成！');
}

// 运行测试
if (typeof module !== 'undefined' && module.exports) {
    // Node.js环境
    module.exports = { runTests, detectMappingMode, parseStatusSheet };
} else {
    // 浏览器环境
    runTests();
}
