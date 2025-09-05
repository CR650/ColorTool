/**
 * 颜色主题管理模块
 * 文件说明：负责Unity项目颜色主题的管理，包括主题数据的映射、更新和生成
 * 创建时间：2025-01-09
 * 依赖：App.Utils, App.DataParser, XLSX库
 */

// 确保全局命名空间存在
window.App = window.App || {};

/**
 * 颜色主题管理模块
 * 处理颜色主题的创建、更新和管理功能
 */
window.App.ThemeManager = (function() {
    'use strict';

    // 模块状态
    let isInitialized = false;
    let sourceData = null;           // 源数据文件内容
    let unityProjectFiles = null;    // Unity项目文件列表
    let rscThemeData = null;         // RSC_Theme.xls文件数据
    let ugcThemeData = null;         // UGCTheme.xls文件数据
    let mappingData = null;          // 对比映射数据
    let processedResult = null;      // 处理结果
    let rscAllSheetsData = null;     // RSC_Theme文件的所有Sheet数据
    let ugcAllSheetsData = null;     // UGCTheme文件的所有Sheet数据

    // DOM元素引用
    let themeNameInput = null;
    let processThemeBtn = null;
    let resetBtn = null;
    let enableDirectSaveBtn = null;
    let enableUGCDirectSaveBtn = null;
    let fileStatus = null;

    // Sheet选择器相关DOM元素
    let sheetSelectorSection = null;
    let fileTypeSelect = null;
    let rscSheetSelect = null;
    let rscSheetInfo = null;
    let sheetDataContainer = null;
    let sheetDataStats = null;
    let rscSheetTable = null;
    let rscSheetTableHead = null;
    let rscSheetTableBody = null;

    /**
     * 初始化主题管理模块
     */
    function init() {
        if (isInitialized) {
            console.warn('ThemeManager模块已经初始化');
            return;
        }

        // 获取DOM元素
        themeNameInput = document.getElementById('themeNameInput');
        processThemeBtn = document.getElementById('processThemeBtn');
        resetBtn = document.getElementById('resetBtn');
        enableDirectSaveBtn = document.getElementById('enableDirectSaveBtn');
        enableUGCDirectSaveBtn = document.getElementById('enableUGCDirectSaveBtn');
        fileStatus = document.getElementById('fileStatus');

        // 获取Sheet选择器相关DOM元素
        sheetSelectorSection = document.getElementById('sheetSelectorSection');
        fileTypeSelect = document.getElementById('fileTypeSelect');
        rscSheetSelect = document.getElementById('rscSheetSelect');
        rscSheetInfo = document.getElementById('rscSheetInfo');
        sheetDataContainer = document.getElementById('sheetDataContainer');
        sheetDataStats = document.getElementById('sheetDataStats');
        rscSheetTable = document.getElementById('rscSheetTable');
        rscSheetTableHead = document.getElementById('rscSheetTableHead');
        rscSheetTableBody = document.getElementById('rscSheetTableBody');

        // 设置事件监听器
        setupEventListeners();

        // 加载对比映射数据
        loadMappingData();

        // 显示浏览器兼容性信息
        displayBrowserCompatibility();

        isInitialized = true;
        console.log('ThemeManager模块初始化完成');
    }

    /**
     * 设置事件监听器
     */
    function setupEventListeners() {
        if (processThemeBtn) {
            processThemeBtn.addEventListener('click', processThemeData);
        }

        if (resetBtn) {
            resetBtn.addEventListener('click', resetAll);
        }



        if (enableDirectSaveBtn) {
            enableDirectSaveBtn.addEventListener('click', enableDirectFileSave);
        }

        if (enableUGCDirectSaveBtn) {
            enableUGCDirectSaveBtn.addEventListener('click', enableUGCDirectFileSave);
        }

        // 设置Sheet选择器事件监听器
        if (fileTypeSelect) {
            fileTypeSelect.addEventListener('change', handleFileTypeChange);
        }

        if (rscSheetSelect) {
            rscSheetSelect.addEventListener('change', handleSheetSelectionChange);
        }
    }

    /**
     * 加载对比映射数据（内置数据）
     */
    function loadMappingData() {
        // 内置对比映射数据，避免网络请求依赖
        mappingData = {
            "sheetName": "Sheet1",
            "exportTime": "2025-09-04T03:13:30.662Z",
            "totalRows": 33,
            "totalColumns": 7,
            "headers": [
                "RC现在的主题通道",
                "作用",
                "",
                "AI工具导出的颜色表格式",
                "颜色代码",
                "作用",
                "和RC联系"
            ],
            "data": [
                {
                    "RC现在的主题通道": "P1",
                    "作用": "地板颜色",
                    "颜色代码": "P1"
                },
                {
                    "RC现在的主题通道": "P5",
                    "作用": "跳板颜色",
                    "颜色代码": "P2"
                },
                {
                    "RC现在的主题通道": "G1",
                    "作用": "装饰颜色1",
                    "颜色代码": "G1"
                },
                {
                    "RC现在的主题通道": "G2",
                    "作用": "装饰颜色2",
                    "颜色代码": "G2"
                },
                {
                    "RC现在的主题通道": "G3",
                    "作用": "装饰颜色3",
                    "颜色代码": "G3"
                },
                {
                    "RC现在的主题通道": "G4",
                    "作用": "装饰颜色4",
                    "颜色代码": "G4"
                },
                {
                    "RC现在的主题通道": "P2",
                    "作用": "地板描边颜色（边框颜色）",
                    "颜色代码": "P1-1"
                },
                {
                    "RC现在的主题通道": "P9",
                    "作用": "地板侧面颜色（立面）",
                    "颜色代码": "P1-2"
                },
                {
                    "RC现在的主题通道": "P6",
                    "作用": "跳板描边颜色（边框颜色）",
                    "颜色代码": "P2-1"
                },
                {
                    "RC现在的主题通道": "P10",
                    "作用": "跳板侧面颜色（立面）",
                    "颜色代码": "P2-2"
                },
                {
                    "RC现在的主题通道": "G5",
                    "作用": "装饰颜色5",
                    "颜色代码": "G5"
                },
                {
                    "RC现在的主题通道": "G6",
                    "作用": "装饰颜色6",
                    "颜色代码": "G6"
                },
                {
                    "RC现在的主题通道": "G7",
                    "作用": "装饰颜色7",
                    "颜色代码": "G7"
                },
                {
                    "RC现在的主题通道": "P3",
                    "作用": "预留颜色通道3",
                    "颜色代码": "P3"
                },
                {
                    "RC现在的主题通道": "P4",
                    "作用": "预留颜色通道4",
                    "颜色代码": "P4"
                },
                {
                    "RC现在的主题通道": "P7",
                    "作用": "预留颜色通道7",
                    "颜色代码": "P7"
                },
                {
                    "RC现在的主题通道": "P8",
                    "作用": "预留颜色通道8",
                    "颜色代码": "P8"
                }
            ]
        };

        updateFileStatus('mappingStatus', '已加载', 'success');
        console.log('对比映射数据加载成功（内置数据）');
    }

    /**
     * 设置源数据
     * @param {Object} data - 源数据文件内容
     */
    function setSourceData(data) {
        sourceData = data;
        updateFileStatus('sourceFileStatus', `已选择: ${data.fileName}`, 'success');

        // 调试：输出源数据结构
        console.log('源数据加载完成:', {
            fileName: data.fileName,
            headers: data.headers,
            dataCount: data.data.length,
            sampleData: data.data.slice(0, 3)
        });

        checkReadyState();
    }

    /**
     * 设置Unity项目文件
     * @param {FileList} files - Unity项目文件列表
     */
    function setUnityProjectFiles(files) {
        unityProjectFiles = files;

        // 查找RSC_Theme.xls文件
        findRSCThemeFile();
        checkReadyState();
    }

    /**
     * 直接设置RSC_Theme文件
     * @param {File} file - RSC_Theme文件对象
     */
    function setRSCThemeFile(file) {
        loadRSCThemeFile(file);
        checkReadyState();
    }

    /**
     * 查找RSC_Theme.xls文件（优化版，异步处理）
     */
    function findRSCThemeFile() {
        if (!unityProjectFiles) {
            return;
        }

        console.log(`开始在${unityProjectFiles.length}个文件中查找RSC_Theme.xls...`);
        App.Utils.showStatus('正在查找RSC_Theme.xls文件...', 'info');

        // 使用异步处理避免界面卡顿
        setTimeout(() => {
            let found = false;

            // 在文件列表中查找RSC_Theme.xls
            for (let i = 0; i < unityProjectFiles.length; i++) {
                const file = unityProjectFiles[i];

                // 每处理100个文件更新一次进度
                if (i % 100 === 0 && i > 0) {
                    App.Utils.showStatus(`正在查找RSC_Theme.xls文件... (${i}/${unityProjectFiles.length})`, 'info');
                }

                if (file.name === 'RSC_Theme.xls' || file.name === 'RSC_Theme.xlsx') {
                    console.log(`找到RSC_Theme文件: ${file.name}, 路径: ${file.webkitRelativePath}`);
                    loadRSCThemeFile(file);
                    found = true;
                    break;
                }
            }

            if (!found) {
                updateFileStatus('rscThemeStatus', '未找到', 'error');
                App.Utils.showStatus('在选择的Unity项目文件夹中未找到RSC_Theme.xls文件，请确认文件夹路径正确', 'warning');
                console.warn('RSC_Theme.xls文件未找到，文件列表:', Array.from(unityProjectFiles).map(f => f.name));
            }
        }, 10); // 短暂延迟，让界面有时间更新
    }

    /**
     * 加载RSC_Theme.xls文件
     * @param {File} file - RSC_Theme文件
     */
    function loadRSCThemeFile(file) {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const workbook = XLSX.read(e.target.result, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                    header: 1, 
                    defval: '',
                    raw: false 
                });
                
                rscThemeData = {
                    workbook: workbook,
                    data: jsonData,
                    fileName: file.name
                };

                // 存储所有Sheet数据
                rscAllSheetsData = {};
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const sheetData = XLSX.utils.sheet_to_json(worksheet, {
                        header: 1,
                        defval: '',
                        raw: false
                    });
                    rscAllSheetsData[sheetName] = sheetData;
                });

                console.log('RSC_Theme所有Sheet数据已存储:', Object.keys(rscAllSheetsData));

                updateFileStatus('rscThemeStatus', '已加载', 'success');
                console.log('RSC_Theme.xls文件加载成功');

                // 关键修复：添加状态检查调用
                checkReadyState();
            } catch (error) {
                console.error('RSC_Theme.xls文件解析失败:', error);
                updateFileStatus('rscThemeStatus', '解析失败', 'error');
                App.Utils.showStatus('RSC_Theme.xls文件解析失败', 'error');
            }
        };

        reader.onerror = function() {
            updateFileStatus('rscThemeStatus', '读取失败', 'error');
            App.Utils.showStatus('RSC_Theme.xls文件读取失败', 'error');
        };

        reader.readAsArrayBuffer(file);
    }

    /**
     * 更新文件状态显示
     * @param {string} elementId - 状态元素ID
     * @param {string} text - 状态文本
     * @param {string} type - 状态类型
     */
    function updateFileStatus(elementId, text, type) {
        const element = document.getElementById(elementId);
        if (element) {
            element.textContent = text;
            element.className = `status-value ${type}`;
        }

        // 显示文件状态区域
        if (fileStatus) {
            fileStatus.style.display = 'block';
        }
    }

    /**
     * 检查是否准备就绪
     */
    function checkReadyState() {
        const isReady = sourceData && rscThemeData && ugcThemeData && mappingData;

        if (isReady) {
            // 显示主题输入区域
            const themeInputSection = document.getElementById('themeInputSection');
            if (themeInputSection) {
                themeInputSection.style.display = 'block';
            }

            App.Utils.showStatus('所有文件已准备就绪，请输入主题名称', 'success');
        }
    }

    /**
     * 处理主题数据
     */
    async function processThemeData() {
        const themeName = themeNameInput ? themeNameInput.value.trim() : '';
        
        if (!themeName) {
            App.Utils.showStatus('请输入主题名称', 'error');
            return;
        }

        if (!sourceData || !rscThemeData || !ugcThemeData || !mappingData) {
            App.Utils.showStatus('请确保所有必要文件已加载', 'error');
            return;
        }

        try {
            console.log('=== 开始处理主题数据 ===');
            console.log('主题名称:', themeName);
            console.log('源数据行数:', sourceData ? sourceData.length : 'null');
            console.log('RSC数据行数:', rscThemeData ? rscThemeData.data.length : 'null');
            console.log('映射数据状态:', mappingData ? 'loaded' : 'null');

            App.Utils.showStatus('正在处理主题数据...', 'info');

            // 执行主题数据处理
            console.log('调用 executeThemeProcessing...');
            const result = executeThemeProcessing(themeName);
            console.log('executeThemeProcessing 返回结果:', result);

            if (result.success) {
                processedResult = result;
                console.log('✅ 主题处理成功，结果已保存');
                displayProcessingResult(result);

                // 处理UGCTheme文件（如果是新增主题）
                const ugcResult = await processUGCTheme(themeName, result.isNewTheme);

                // 直接保存文件
                await handleFileSave(result.workbook, result.themeName, ugcResult);
            } else {
                App.Utils.showStatus('主题数据处理失败: ' + result.error, 'error');
            }
        } catch (error) {
            console.error('主题数据处理错误:', error);
            App.Utils.showStatus('主题数据处理失败: ' + error.message, 'error');
        }
    }

    /**
     * 执行主题处理逻辑
     * @param {string} themeName - 主题名称
     * @returns {Object} 处理结果
     */
    function executeThemeProcessing(themeName) {
        try {
            console.log('=== executeThemeProcessing 开始 ===');
            console.log('输入主题名称:', themeName);

            // 保存处理前的数据快照
            const beforeProcessing = {
                dataLength: rscThemeData.data.length,
                lastRow: rscThemeData.data[rscThemeData.data.length - 1]
            };
            console.log('处理前数据快照:', beforeProcessing);

            // 1. 查找或创建主题行
            console.log('步骤1: 查找或创建主题行...');
            const themeRowResult = findOrCreateThemeRow(themeName);
            const themeRowIndex = themeRowResult.index;
            const isNewTheme = themeRowResult.isNew;
            console.log('主题行索引:', themeRowIndex, '是否新主题:', isNewTheme);

            // 2. 根据映射关系更新颜色数据
            console.log('步骤2: 更新颜色数据...');
            const updateResult = updateThemeColors(themeRowIndex, themeName);
            console.log('颜色更新结果:', updateResult);

            // 3. 验证颜色通道处理完整性
            console.log('步骤3: 验证颜色通道完整性...');
            const validationPassed = validateColorChannelCompleteness(updateResult.updatedColors, themeRowIndex);
            console.log('验证结果:', validationPassed);

            // 4. 处理RSC_Theme中的ColorInfo和Light sheet（如果是新增主题）
            console.log('步骤4: 处理RSC_Theme额外sheet...');
            const rscAdditionalResult = processRSCAdditionalSheets(themeName, isNewTheme);
            console.log('RSC额外sheet处理结果:', rscAdditionalResult);

            // 5. 生成更新后的Excel文件
            console.log('步骤5: 生成更新后的Excel文件...');
            const updatedWorkbook = generateUpdatedWorkbook();
            console.log('工作簿生成完成:', updatedWorkbook ? '成功' : '失败');

            // 5. 输出完整的处理结果和数据状态
            console.log('=== 主题处理完成，输出最终数据状态 ===');
            console.log(`处理的主题: ${themeName}`);
            console.log(`主题行索引: ${themeRowIndex}`);
            console.log(`RSC数据总行数: ${rscThemeData.data.length}`);

            // 输出处理后的完整行数据
            const finalThemeRow = rscThemeData.data[themeRowIndex];
            console.log('=== 最终主题行数据 ===');
            console.log('表头:', rscThemeData.data[0]);
            console.log(`行${themeRowIndex}数据:`, finalThemeRow);

            // 输出所有颜色通道的最终值
            console.log('=== 所有颜色通道最终值 ===');
            const headerRow = rscThemeData.data[0];
            updateResult.updatedColors.forEach(colorInfo => {
                const finalValue = finalThemeRow[colorInfo.columnIndex];
                const status = colorInfo.isDefault ? '(默认值)' : '(源数据)';
                console.log(`${colorInfo.channel}: ${finalValue} ${status}`);
            });

            // 验证数据完整性
            const dataIntegrityCheck = {
                themeRowExists: !!finalThemeRow,
                themeRowLength: finalThemeRow ? finalThemeRow.length : 0,
                expectedLength: headerRow.length,
                allColorChannelsHaveValues: updateResult.updatedColors.every(c => {
                    const value = finalThemeRow[c.columnIndex];
                    return value && value !== '';
                })
            };

            console.log('=== 数据完整性检查 ===');
            console.log(dataIntegrityCheck);

            if (dataIntegrityCheck.allColorChannelsHaveValues) {
                console.log('✅ 所有颜色通道都有有效值');
            } else {
                console.error('❌ 存在空的颜色通道值');
            }

            console.log('=== 主题处理和数据状态输出完成 ===');

            // 同步更新内存中的数据状态
            console.log('=== 开始同步内存数据状态 ===');
            syncMemoryDataState(updatedWorkbook, themeRowIndex);

            return {
                success: true,
                themeName: themeName,
                rowIndex: themeRowIndex,
                isNewTheme: isNewTheme,
                updatedColors: updateResult.updatedColors,
                workbook: updatedWorkbook,
                summary: updateResult.summary,
                validationPassed: validationPassed,
                finalThemeRowData: finalThemeRow,
                dataIntegrityCheck: dataIntegrityCheck,
                rscAdditionalSheets: rscAdditionalResult
            };
        } catch (error) {
            return {
                success: false,
                error: error.message
            };
        }
    }

    /**
     * 查找或创建主题行
     * @param {string} themeName - 主题名称
     * @returns {number} 主题行索引
     */
    function findOrCreateThemeRow(themeName) {
        console.log('=== 开始查找或创建主题行 ===');
        console.log(`目标主题名称: ${themeName}`);

        const data = rscThemeData.data;
        console.log(`当前RSC数据行数: ${data.length}`);

        // 查找notes列的索引
        const headerRow = data[0];
        console.log('表头行:', headerRow);
        const notesColumnIndex = headerRow.findIndex(col => col === 'notes');

        if (notesColumnIndex === -1) {
            throw new Error('在RSC_Theme.xls中未找到notes列');
        }
        console.log(`notes列索引: ${notesColumnIndex}`);

        // 查找是否存在相同主题名称的行
        for (let i = 1; i < data.length; i++) {
            if (data[i][notesColumnIndex] === themeName) {
                console.log(`✅ 找到现有主题: ${themeName}, 行索引: ${i}`);
                console.log(`现有行数据:`, data[i]);
                console.log('=== 主题行查找完成（使用现有行） ===');
                return { index: i, isNew: false };
            }
        }

        // 如果没有找到，创建新行
        console.log(`未找到现有主题，开始创建新行...`);
        const newRowIndex = data.length;
        const newRow = new Array(headerRow.length).fill('');

        console.log(`新行索引: ${newRowIndex}`);
        console.log(`新行长度: ${newRow.length} (表头长度: ${headerRow.length})`);

        // 设置id字段（自动递增）
        const idColumnIndex = headerRow.findIndex(col => col === 'id');
        if (idColumnIndex !== -1) {
            const existingIds = data.slice(1).map(row => parseInt(row[idColumnIndex]) || 0);
            const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 0;
            const newId = maxId + 1;
            newRow[idColumnIndex] = newId.toString();
            console.log(`设置ID字段: 列${idColumnIndex} = ${newId}`);
        }

        // 设置notes字段
        newRow[notesColumnIndex] = themeName;
        console.log(`设置notes字段: 列${notesColumnIndex} = ${themeName}`);

        // 添加新行到数据数组
        data.push(newRow);
        console.log(`✅ 新行已添加到数据数组，当前总行数: ${data.length}`);

        // 验证新行是否正确添加
        const addedRow = data[newRowIndex];
        if (addedRow === newRow) {
            console.log(`✅ 数据引用验证通过: data[${newRowIndex}] === newRow`);
        } else {
            console.error(`❌ 数据引用验证失败: data[${newRowIndex}] !== newRow`);
        }

        console.log(`新创建的行数据:`, addedRow);
        console.log('=== 主题行创建完成 ===');

        return { index: newRowIndex, isNew: true };
    }

    /**
     * 更新主题颜色数据
     * @param {number} rowIndex - 主题行索引
     * @param {string} themeName - 主题名称
     * @returns {Object} 更新结果
     */
    function updateThemeColors(rowIndex, themeName) {
        console.log('=== 开始更新主题颜色数据 ===');
        console.log(`目标行索引: ${rowIndex}, 主题名称: ${themeName}`);

        const data = rscThemeData.data;
        const headerRow = data[0];
        const themeRow = data[rowIndex];

        // 验证数据引用的正确性
        console.log(`RSC数据总行数: ${data.length}`);
        console.log(`目标行索引: ${rowIndex}`);

        if (!themeRow) {
            throw new Error(`无效的行索引: ${rowIndex}, 数据总行数: ${data.length}`);
        }

        console.log(`✅ 目标行数据引用获取成功`);
        console.log(`目标行当前数据:`, themeRow);
        console.log(`目标行数据长度: ${themeRow.length}, 表头长度: ${headerRow.length}`);

        // 验证这是同一个对象引用
        if (data[rowIndex] === themeRow) {
            console.log(`✅ 数据引用验证通过: data[${rowIndex}] === themeRow`);
        } else {
            console.error(`❌ 数据引用验证失败: data[${rowIndex}] !== themeRow`);
        }

        const updatedColors = [];
        const summary = {
            total: 0,
            updated: 0,
            notFound: 0,
            errors: []
        };

        console.log('映射数据:', mappingData.data);
        console.log('RSC_Theme表头:', headerRow);

        // 遍历映射数据
        mappingData.data.forEach((mapping, index) => {
            const rcChannel = mapping['RC现在的主题通道'];
            const colorCode = mapping['颜色代码'];

            console.log(`\n处理映射 ${index + 1}:`, {
                rcChannel,
                colorCode,
                作用: mapping['作用']
            });

            // 跳过空的RC通道或标记为"占不导入"的项目
            if (!rcChannel || rcChannel === '占不导入' || rcChannel === '' || rcChannel === '暂不导入') {
                console.log(`跳过映射: RC通道为空或标记为不导入 (${rcChannel})`);
                return;
            }

            summary.total++;

            try {
                // 在RSC_Theme表头中查找对应的列
                const columnIndex = headerRow.findIndex(col => col === rcChannel);

                if (columnIndex === -1) {
                    const error = `未找到列: ${rcChannel}`;
                    console.error(error);
                    summary.errors.push(error);
                    return;
                }

                console.log(`找到目标列: ${rcChannel} (索引: ${columnIndex})`);

                // 从源数据中查找对应的颜色值
                const colorValue = findColorValue(colorCode);

                // 确保颜色值处理的健壮性
                let finalColorValue = null;
                let isDefault = false;

                if (colorValue && colorValue !== null && colorValue !== undefined && colorValue !== '') {
                    // 验证颜色值格式
                    const cleanColorValue = colorValue.toString().trim().toUpperCase();
                    if (/^[0-9A-F]{6}$/i.test(cleanColorValue)) {
                        finalColorValue = cleanColorValue;
                    } else {
                        console.warn(`⚠️ 颜色值格式无效: ${colorValue}, 使用默认值`);
                        finalColorValue = 'FFFFFF';
                        isDefault = true;
                    }
                } else {
                    // 使用默认值FFFFFF
                    finalColorValue = 'FFFFFF';
                    isDefault = true;
                }

                // 确保数据更新到正确的位置
                if (themeRow && columnIndex >= 0 && columnIndex < themeRow.length) {
                    themeRow[columnIndex] = finalColorValue;
                    console.log(`📝 数据更新: 行${rowIndex}, 列${columnIndex}(${rcChannel}) = ${finalColorValue}`);
                } else {
                    console.error(`❌ 数据更新失败: 无效的行或列索引 - 行:${rowIndex}, 列:${columnIndex}`);
                    throw new Error(`无效的数据位置: 行${rowIndex}, 列${columnIndex}`);
                }

                // 记录更新结果
                updatedColors.push({
                    channel: rcChannel,
                    colorCode: colorCode,
                    value: finalColorValue,
                    isDefault: isDefault,
                    rowIndex: rowIndex,
                    columnIndex: columnIndex
                });

                if (isDefault) {
                    summary.notFound++;
                    console.warn(`⚠️ 使用默认值: ${rcChannel} = ${finalColorValue} (颜色代码: ${colorCode})`);
                } else {
                    summary.updated++;
                    console.log(`✅ 成功更新: ${rcChannel} = ${finalColorValue} (颜色代码: ${colorCode})`);
                }
            } catch (error) {
                const errorMsg = `处理${rcChannel}时出错: ${error.message}`;
                console.error(errorMsg, error);
                summary.errors.push(errorMsg);
            }
        });

        console.log('\n颜色映射处理完成:', summary);

        // 处理所有颜色通道，确保没有映射的通道也有默认值
        processAllColorChannels(headerRow, themeRow, rowIndex, updatedColors, summary);

        // 验证数据更新结果
        console.log('=== 数据更新验证 ===');
        console.log(`主题行数据 (行${rowIndex}):`, themeRow);

        // 验证数据引用一致性
        const dataRowReference = rscThemeData.data[rowIndex];
        if (dataRowReference === themeRow) {
            console.log(`✅ 数据引用一致性验证通过: rscThemeData.data[${rowIndex}] === themeRow`);
        } else {
            console.error(`❌ 数据引用一致性验证失败: rscThemeData.data[${rowIndex}] !== themeRow`);
            console.log('rscThemeData.data[rowIndex]:', dataRowReference);
            console.log('themeRow:', themeRow);
        }

        // 验证每个更新的颜色通道
        let verificationErrors = 0;
        updatedColors.forEach(colorInfo => {
            const actualValueInThemeRow = themeRow[colorInfo.columnIndex];
            const actualValueInDataArray = rscThemeData.data[rowIndex][colorInfo.columnIndex];

            if (actualValueInThemeRow === colorInfo.value && actualValueInDataArray === colorInfo.value) {
                console.log(`✅ 验证通过: ${colorInfo.channel} = ${actualValueInThemeRow}`);
            } else {
                console.error(`❌ 验证失败: ${colorInfo.channel}`);
                console.error(`  期望值: ${colorInfo.value}`);
                console.error(`  themeRow中的值: ${actualValueInThemeRow}`);
                console.error(`  data数组中的值: ${actualValueInDataArray}`);
                verificationErrors++;
            }
        });

        // 输出完整的更新后行数据
        console.log('=== 完整的更新后行数据 ===');
        console.log(`行索引: ${rowIndex}`);
        console.log('表头:', headerRow);
        console.log('数据:', themeRow);
        console.log('数据数组中的同一行:', rscThemeData.data[rowIndex]);

        // 验证关键字段
        const notesColumnIndex = headerRow.findIndex(col => col === 'notes');
        if (notesColumnIndex !== -1) {
            console.log(`notes字段值: ${themeRow[notesColumnIndex]}`);
        }

        if (verificationErrors === 0) {
            console.log('✅ 所有颜色通道验证通过');
        } else {
            console.error(`❌ ${verificationErrors}个颜色通道验证失败`);
        }

        console.log('=== 数据更新验证完成 ===');

        return {
            updatedColors: updatedColors,
            summary: summary,
            themeRowData: themeRow,
            verificationPassed: true
        };
    }

    /**
     * 处理所有颜色通道，确保没有映射的通道也有默认值
     * @param {Array} headerRow - 表头行
     * @param {Array} themeRow - 主题行数据
     * @param {number} rowIndex - 行索引
     * @param {Array} updatedColors - 已更新的颜色列表
     * @param {Object} summary - 处理摘要
     */
    function processAllColorChannels(headerRow, themeRow, rowIndex, updatedColors, summary) {
        console.log('\n=== 开始处理所有颜色通道 ===');

        // 识别所有颜色通道列（P开头和G开头的列）
        const colorChannels = headerRow.filter((col) => {
            if (!col || typeof col !== 'string') return false;
            const colName = col.toString().trim().toUpperCase();
            return colName.startsWith('P') || colName.startsWith('G');
        });

        console.log('发现的颜色通道:', colorChannels);

        // 检查每个颜色通道是否已经被处理
        colorChannels.forEach(channel => {
            const columnIndex = headerRow.findIndex(col => col === channel);
            const alreadyProcessed = updatedColors.find(c => c.channel === channel);

            if (!alreadyProcessed) {
                console.log(`处理未映射的颜色通道: ${channel}`);

                // 检查当前值是否为空或无效
                const currentValue = themeRow[columnIndex];
                let needsDefault = false;

                if (!currentValue || currentValue === '' || currentValue === null || currentValue === undefined) {
                    needsDefault = true;
                } else {
                    // 检查是否为有效的颜色值
                    const cleanValue = currentValue.toString().trim().toUpperCase();
                    if (!/^[0-9A-F]{6}$/i.test(cleanValue)) {
                        needsDefault = true;
                    }
                }

                if (needsDefault) {
                    // 设置默认值
                    const defaultValue = 'FFFFFF';
                    themeRow[columnIndex] = defaultValue;

                    // 记录更新结果
                    updatedColors.push({
                        channel: channel,
                        colorCode: '无映射',
                        value: defaultValue,
                        isDefault: true,
                        rowIndex: rowIndex,
                        columnIndex: columnIndex
                    });

                    summary.notFound++;
                    console.log(`✅ 设置默认值: ${channel} = ${defaultValue} (无映射)`);
                } else {
                    console.log(`✅ 保持现有值: ${channel} = ${currentValue}`);
                }
            } else {
                console.log(`✅ 已处理: ${channel} = ${alreadyProcessed.value}`);
            }
        });

        console.log('=== 所有颜色通道处理完成 ===\n');
    }

    /**
     * 处理RSC_Theme文件中的ColorInfo和Light sheet（新增主题时添加新行）
     * @param {string} themeName - 主题名称
     * @param {boolean} isNewTheme - 是否为新增主题
     * @returns {Object} 处理结果
     */
    function processRSCAdditionalSheets(themeName, isNewTheme) {
        console.log('=== 开始处理RSC_Theme的ColorInfo和Light sheet ===');
        console.log('主题名称:', themeName);
        console.log('是否新增主题:', isNewTheme);

        if (!isNewTheme) {
            console.log('覆盖现有主题，不需要处理ColorInfo和Light sheet');
            return { success: true, action: 'skip', reason: '覆盖现有主题' };
        }

        if (!rscThemeData || !rscThemeData.workbook) {
            console.error('RSC_Theme数据未加载');
            return { success: false, error: 'RSC_Theme数据未加载' };
        }

        try {
            const workbook = rscThemeData.workbook;
            const sheetNames = workbook.SheetNames;
            console.log('RSC_Theme包含的sheet:', sheetNames);

            const targetSheets = ['ColorInfo', 'Light'];
            const processedSheets = [];

            targetSheets.forEach(sheetName => {
                if (sheetNames.includes(sheetName)) {
                    console.log(`开始处理sheet: ${sheetName}`);

                    const worksheet = workbook.Sheets[sheetName];
                    const sheetData = XLSX.utils.sheet_to_json(worksheet, {
                        header: 1,
                        defval: '',
                        raw: false
                    });

                    if (sheetData.length > 0) {
                        const result = addNewRowToSheet(sheetData, themeName, sheetName);
                        if (result.success) {
                            // 更新工作表
                            const newWorksheet = XLSX.utils.aoa_to_sheet(sheetData);
                            workbook.Sheets[sheetName] = newWorksheet;

                            processedSheets.push({
                                sheetName: sheetName,
                                newRowIndex: result.newRowIndex,
                                newId: result.newId
                            });
                            console.log(`✅ ${sheetName} sheet处理成功`);
                        } else {
                            console.warn(`⚠️ ${sheetName} sheet处理失败:`, result.error);
                        }
                    } else {
                        console.warn(`⚠️ ${sheetName} sheet为空，跳过处理`);
                    }
                } else {
                    console.log(`Sheet "${sheetName}" 不存在，跳过处理`);
                }
            });

            console.log('RSC_Theme额外sheet处理完成，处理的sheets:', processedSheets);

            return {
                success: true,
                action: 'add_new_rows',
                processedSheets: processedSheets,
                workbook: workbook
            };

        } catch (error) {
            console.error('处理RSC_Theme额外sheet失败:', error);
            return {
                success: false,
                error: error.message
            };
        }
    }

    /**
     * 向指定sheet添加新行
     * @param {Array} sheetData - sheet数据数组
     * @param {string} themeName - 主题名称
     * @param {string} sheetName - sheet名称
     * @returns {Object} 处理结果
     */
    function addNewRowToSheet(sheetData, themeName, sheetName) {
        console.log(`=== 开始向${sheetName}添加新行 ===`);

        if (sheetData.length === 0) {
            return { success: false, error: 'Sheet数据为空' };
        }

        const headerRow = sheetData[0];
        console.log(`${sheetName} 表头:`, headerRow);

        // 查找id和notes列的索引
        const idColumnIndex = headerRow.findIndex(col => col === 'id');
        const notesColumnIndex = headerRow.findIndex(col => col === 'notes');

        if (idColumnIndex === -1) {
            console.warn(`${sheetName} 中未找到id列`);
        }
        if (notesColumnIndex === -1) {
            console.warn(`${sheetName} 中未找到notes列`);
        }

        // 创建新行，复制最后一行的数据作为模板
        const lastDataRowIndex = sheetData.length - 1;
        const templateRow = sheetData[lastDataRowIndex];
        const newRow = [...templateRow]; // 复制最后一行数据

        console.log(`使用第${lastDataRowIndex}行作为模板:`, templateRow);

        // 设置id字段（自动递增）
        let newId = null;
        if (idColumnIndex !== -1) {
            const existingIds = sheetData.slice(1).map(row => parseInt(row[idColumnIndex]) || 0);
            const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 0;
            newId = maxId + 1;
            newRow[idColumnIndex] = newId.toString();
            console.log(`设置ID字段: 列${idColumnIndex} = ${newId}`);
        }

        // 设置notes字段
        if (notesColumnIndex !== -1) {
            newRow[notesColumnIndex] = themeName;
            console.log(`设置notes字段: 列${notesColumnIndex} = ${themeName}`);
        }

        // 添加新行到数据数组
        const newRowIndex = sheetData.length;
        sheetData.push(newRow);
        console.log(`✅ 新行已添加到${sheetName}，行索引: ${newRowIndex}`);
        console.log(`新行数据:`, newRow);

        return {
            success: true,
            newRowIndex: newRowIndex,
            newId: newId,
            newRow: newRow
        };
    }

    /**
     * 处理UGCTheme文件（新增主题时添加新行）
     * @param {string} themeName - 主题名称
     * @param {boolean} isNewTheme - 是否为新增主题
     * @returns {Object} 处理结果
     */
    async function processUGCTheme(themeName, isNewTheme) {
        console.log('=== 开始处理UGCTheme文件 ===');
        console.log('主题名称:', themeName);
        console.log('是否新增主题:', isNewTheme);

        if (!isNewTheme) {
            console.log('覆盖现有主题，暂时不处理UGCTheme');
            return { success: true, action: 'skip', reason: '覆盖现有主题' };
        }

        if (!ugcThemeData || !ugcThemeData.workbook) {
            console.error('UGCTheme数据未加载');
            return { success: false, error: 'UGCTheme数据未加载' };
        }

        try {
            const workbook = ugcThemeData.workbook;
            const sheetNames = workbook.SheetNames;
            console.log('UGCTheme包含的sheet:', sheetNames);

            const processedSheets = [];

            // 对每个sheet进行处理
            for (const sheetName of sheetNames) {
                console.log(`处理sheet: ${sheetName}`);
                const worksheet = workbook.Sheets[sheetName];
                const data = XLSX.utils.sheet_to_json(worksheet, {
                    header: 1,
                    defval: '',
                    raw: false
                });

                if (data.length === 0) {
                    console.log(`Sheet ${sheetName} 为空，跳过`);
                    continue;
                }

                const headerRow = data[0];
                console.log(`Sheet ${sheetName} 表头:`, headerRow);

                // 查找id列
                const idColumnIndex = headerRow.findIndex(col => col === 'id');
                if (idColumnIndex === -1) {
                    console.log(`Sheet ${sheetName} 没有id列，跳过`);
                    continue;
                }

                // 获取最大ID
                const existingIds = data.slice(1).map(row => parseInt(row[idColumnIndex]) || 0);
                const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 0;
                const newId = maxId + 1;

                // 创建新行（复制上一行数据）
                const lastRow = data[data.length - 1];
                const newRow = [...lastRow]; // 复制上一行
                newRow[idColumnIndex] = newId.toString(); // 设置新的ID

                console.log(`Sheet ${sheetName} 新行:`, newRow);

                // 添加新行到数据
                data.push(newRow);

                // 更新worksheet
                const newWorksheet = XLSX.utils.aoa_to_sheet(data);
                workbook.Sheets[sheetName] = newWorksheet;

                processedSheets.push({
                    sheetName: sheetName,
                    newId: newId,
                    newRowIndex: data.length - 1
                });
            }

            console.log('UGCTheme处理完成，处理的sheets:', processedSheets);

            // 同步UGC内存数据状态
            console.log('=== 开始同步UGC内存数据状态 ===');
            syncUGCMemoryDataState(workbook);

            return {
                success: true,
                action: 'add_new_rows',
                processedSheets: processedSheets,
                workbook: workbook
            };

        } catch (error) {
            console.error('处理UGCTheme失败:', error);
            return {
                success: false,
                error: error.message
            };
        }
    }

    /**
     * 从源数据中查找颜色值（增强版，包含详细调试）
     * @param {string} colorCode - 颜色代码
     * @returns {string|null} 16进制颜色值
     */
    function findColorValue(colorCode) {
        if (!sourceData || !sourceData.data) {
            console.warn(`查找颜色值失败: 源数据未加载, colorCode=${colorCode}`);
            return null;
        }

        console.log(`开始查找颜色代码: ${colorCode}`);
        console.log(`源数据包含 ${sourceData.data.length} 行数据`);
        console.log(`源数据表头:`, sourceData.headers);
        console.log('源数据验证信息:', sourceData.validation);

        // 在源数据中查找匹配的颜色代码
        for (let i = 0; i < sourceData.data.length; i++) {
            const row = sourceData.data[i];

            // 调试：输出每行的颜色代码字段
            if (i < 5) { // 只输出前5行避免日志过多
                console.log(`第${i}行数据:`, {
                    颜色代码: row['颜色代码'],
                    '16进制值': row['16进制值'],
                    R值: row['R值'],
                    G值: row['G值'],
                    B值: row['B值']
                });
            }

            // 多种匹配策略：精确匹配和模糊匹配
            const colorCodeFields = ['颜色代码', 'colorCode', 'code', '代码', 'Color Code', 'ColorCode'];
            let rowColorCode = null;

            // 尝试从多个可能的字段中获取颜色代码
            for (const field of colorCodeFields) {
                if (row[field] && row[field] !== '') {
                    rowColorCode = row[field];
                    break;
                }
            }

            // 如果没有找到标准字段，尝试使用第一个非空字段
            if (!rowColorCode) {
                const firstNonEmptyValue = Object.values(row).find(value => value && value !== '');
                if (firstNonEmptyValue) {
                    rowColorCode = firstNonEmptyValue;
                    console.log(`使用第一个非空值作为颜色代码: ${firstNonEmptyValue}`);
                }
            }

            if (rowColorCode && (rowColorCode === colorCode ||
                                rowColorCode === colorCode.toUpperCase() ||
                                rowColorCode === colorCode.toLowerCase() ||
                                rowColorCode.toString().trim() === colorCode.toString().trim())) {
                console.log(`找到匹配的颜色代码 ${colorCode}:`, row);

                // 多种16进制值字段匹配策略
                const hexFields = ['16进制值', '颜色值', 'hex', 'HEX', 'hexValue', '16进制', 'color', 'Color'];
                let hexValue = null;

                for (const field of hexFields) {
                    if (row[field] && row[field] !== '') {
                        hexValue = row[field];
                        console.log(`从字段"${field}"获取16进制值: ${hexValue}`);
                        break;
                    }
                }

                if (hexValue) {
                    console.log(`原始16进制值: ${hexValue}`);

                    // 清理16进制值（移除#号、空格、rgb()等）
                    let cleanHex = hexValue.toString().trim();

                    // 处理rgb(r,g,b)格式
                    const rgbMatch = cleanHex.match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
                    if (rgbMatch) {
                        const r = parseInt(rgbMatch[1]);
                        const g = parseInt(rgbMatch[2]);
                        const b = parseInt(rgbMatch[3]);
                        cleanHex = ((r << 16) | (g << 8) | b).toString(16).padStart(6, '0').toUpperCase();
                        console.log(`从RGB格式转换: rgb(${r},${g},${b}) → ${cleanHex}`);
                    } else {
                        // 移除#号和其他字符
                        cleanHex = cleanHex.replace(/^#/, '').replace(/[^0-9A-Fa-f]/g, '').toUpperCase();
                        console.log(`清理后16进制值: ${cleanHex}`);
                    }

                    // 验证是否为有效的16进制颜色值
                    if (/^[0-9A-F]{6}$/i.test(cleanHex)) {
                        console.log(`✅ 成功提取颜色值: ${colorCode} → ${cleanHex}`);
                        return cleanHex;
                    } else {
                        console.warn(`⚠️ 16进制值格式无效: ${cleanHex} (原始值: ${hexValue})`);
                    }
                }

                // 如果没有16进制值，尝试从RGB值构建
                const rgbFields = [
                    ['R值', 'G值', 'B值'],
                    ['R', 'G', 'B'],
                    ['r', 'g', 'b'],
                    ['red', 'green', 'blue'],
                    ['Red', 'Green', 'Blue']
                ];

                let r, g, b;
                for (const fieldSet of rgbFields) {
                    const [rField, gField, bField] = fieldSet;
                    if (row[rField] !== undefined && row[gField] !== undefined && row[bField] !== undefined) {
                        r = row[rField];
                        g = row[gField];
                        b = row[bField];
                        console.log(`从字段组[${fieldSet.join(', ')}]获取RGB值: R=${r}, G=${g}, B=${b}`);
                        break;
                    }
                }

                if (r !== undefined && g !== undefined && b !== undefined) {
                    const rNum = parseInt(r);
                    const gNum = parseInt(g);
                    const bNum = parseInt(b);

                    console.log(`解析RGB值: R=${rNum}, G=${gNum}, B=${bNum}`);

                    if (!isNaN(rNum) && !isNaN(gNum) && !isNaN(bNum) &&
                        rNum >= 0 && rNum <= 255 && gNum >= 0 && gNum <= 255 && bNum >= 0 && bNum <= 255) {

                        const hexR = rNum.toString(16).padStart(2, '0');
                        const hexG = gNum.toString(16).padStart(2, '0');
                        const hexB = bNum.toString(16).padStart(2, '0');
                        const hexColor = (hexR + hexG + hexB).toUpperCase();

                        console.log(`✅ 从RGB构建颜色值: ${colorCode} → RGB(${rNum},${gNum},${bNum}) → ${hexColor}`);
                        return hexColor;
                    } else {
                        console.warn(`⚠️ RGB值无效: R=${rNum}, G=${gNum}, B=${bNum}`);
                    }
                } else {
                    console.warn(`⚠️ 未找到有效的RGB字段`);
                }

                console.warn(`❌ 无法从行数据中提取有效颜色值:`, row);
                return null;
            }
        }

        console.warn(`❌ 未找到颜色代码: ${colorCode}`);
        return null;
    }

    /**
     * 生成更新后的Excel工作簿
     * @returns {Object} 更新后的工作簿
     */
    function generateUpdatedWorkbook() {
        console.log('=== 开始生成更新后的Excel工作簿 ===');

        if (!rscThemeData || !rscThemeData.workbook || !rscThemeData.data) {
            throw new Error('RSC主题数据不完整，无法生成工作簿');
        }

        const workbook = rscThemeData.workbook;
        const originalSheetName = workbook.SheetNames[0];

        console.log(`原始工作表名称: ${originalSheetName}`);
        console.log(`数据行数: ${rscThemeData.data.length}`);
        console.log(`数据列数: ${rscThemeData.data[0] ? rscThemeData.data[0].length : 0}`);

        // 详细的数据完整性检查
        console.log('=== 详细数据完整性检查 ===');
        if (!rscThemeData.data || rscThemeData.data.length === 0) {
            throw new Error('RSC主题数据为空，无法生成工作簿');
        }

        // 检查数据结构
        const dataIntegrityCheck = {
            totalRows: rscThemeData.data.length,
            headerRow: rscThemeData.data[0],
            lastRow: rscThemeData.data[rscThemeData.data.length - 1],
            hasEmptyRows: rscThemeData.data.some((row, index) => {
                if (!row || row.length === 0) {
                    console.warn(`发现空行在索引 ${index}`);
                    return true;
                }
                return false;
            }),
            maxColumnCount: Math.max(...rscThemeData.data.map(row => row ? row.length : 0))
        };

        console.log('数据完整性检查结果:', dataIntegrityCheck);
        console.log('表头行:', dataIntegrityCheck.headerRow);
        console.log('最后一行:', dataIntegrityCheck.lastRow);

        if (dataIntegrityCheck.hasEmptyRows) {
            console.warn('⚠️ 发现空行，这可能导致Excel文件问题');
        }

        // 将更新后的数据写回工作表
        const newWorksheet = XLSX.utils.aoa_to_sheet(rscThemeData.data);

        // 保持原有的工作表属性（如果有的话）
        const originalWorksheet = workbook.Sheets[originalSheetName];
        if (originalWorksheet && originalWorksheet['!ref']) {
            // 更新工作表范围
            newWorksheet['!ref'] = XLSX.utils.encode_range({
                s: { c: 0, r: 0 },
                e: {
                    c: rscThemeData.data[0].length - 1,
                    r: rscThemeData.data.length - 1
                }
            });
        }

        // 替换工作表
        workbook.Sheets[originalSheetName] = newWorksheet;

        console.log('✅ Excel工作簿生成完成');
        console.log('=== Excel工作簿生成完成 ===');

        return workbook;
    }

    /**
     * 显示处理结果
     * @param {Object} result - 处理结果
     */
    function displayProcessingResult(result) {
        const resultDisplay = document.getElementById('resultDisplay');
        const resultSummary = document.getElementById('resultSummary');
        const resultDetails = document.getElementById('resultDetails');

        if (!resultDisplay || !resultSummary || !resultDetails) {
            return;
        }

        // 显示结果区域
        resultDisplay.style.display = 'block';

        // 生成摘要信息
        const summary = result.summary;
        const validationStatus = result.validationPassed ? '✅ 通过' : '❌ 失败';
        const dataIntegrityStatus = result.dataIntegrityCheck?.allColorChannelsHaveValues ? '✅ 完整' : '❌ 不完整';

        // 处理额外sheet信息
        let additionalSheetsInfo = '';
        if (result.rscAdditionalSheets && result.rscAdditionalSheets.success && result.rscAdditionalSheets.processedSheets) {
            const processedSheets = result.rscAdditionalSheets.processedSheets;
            if (processedSheets.length > 0) {
                additionalSheetsInfo = `
                    <p><strong>额外Sheet处理:</strong> ${processedSheets.map(sheet =>
                        `${sheet.sheetName}(ID:${sheet.newId})`
                    ).join(', ')}</p>
                `;
            }
        }

        resultSummary.innerHTML = `
            <h4>🎯 处理摘要</h4>
            <div style="background-color: #e8f5e8; border: 1px solid #4caf50; padding: 15px; border-radius: 5px; margin: 10px 0;">
                <p><strong>主题名称:</strong> ${result.themeName}</p>
                <p><strong>处理行索引:</strong> ${result.rowIndex}</p>
                <p><strong>是否新增主题:</strong> ${result.isNewTheme ? '✅ 是' : '❌ 否'}</p>
                <p><strong>总计处理:</strong> ${summary.total} 个颜色通道</p>
                <p><strong>成功更新:</strong> ${summary.updated} 个</p>
                <p><strong>使用默认值:</strong> ${summary.notFound} 个</p>
                <p><strong>完整性验证:</strong> ${validationStatus}</p>
                <p><strong>数据完整性:</strong> ${dataIntegrityStatus}</p>
                ${additionalSheetsInfo}
                ${summary.errors.length > 0 ? `<p><strong>错误:</strong> ${summary.errors.length} 个</p>` : ''}
            </div>

        `;

        // 生成详细信息
        let detailsHtml = '<h4>详细信息</h4>\n';
        result.updatedColors.forEach(color => {
            const status = color.isDefault ? '(默认值)' : '';
            detailsHtml += `${color.channel}: ${color.colorCode} → #${color.value} ${status}\n`;
        });

        if (summary.errors.length > 0) {
            detailsHtml += '\n错误信息:\n';
            summary.errors.forEach(error => {
                detailsHtml += `- ${error}\n`;
            });
        }

        resultDetails.textContent = detailsHtml;

        // 自动初始化Sheet选择器
        setTimeout(() => {
            initializeSheetSelector();
        }, 100);

    }

    /**
     * 处理文件保存（整合直接保存和传统下载）
     * @param {Object} workbook - 更新后的RSC工作簿
     * @param {string} themeName - 主题名称
     * @param {Object} ugcResult - UGC处理结果
     */
    async function handleFileSave(workbook, themeName, ugcResult) {
        try {
            // 检查是否支持直接保存
            if (rscThemeData && rscThemeData.fileHandle) {
                const saveDirectly = await confirmDirectSave();
                if (saveDirectly) {
                    // 保存RSC_Theme文件
                    const rscSuccess = await saveFileDirectly(workbook);

                    // 保存UGCTheme文件（如果有处理结果）
                    let ugcSuccess = true;
                    let ugcMessage = '';

                    if (ugcResult && ugcResult.success && ugcResult.workbook && ugcThemeData && ugcThemeData.fileHandle) {
                        console.log('开始保存UGCTheme文件...');
                        ugcSuccess = await saveUGCFileDirectly(ugcResult.workbook);
                        ugcMessage = ugcSuccess ? 'UGCTheme文件保存成功' : 'UGCTheme文件保存失败';
                    } else if (ugcResult && ugcResult.success) {
                        ugcMessage = 'UGCTheme文件未选择或无需处理';
                    }

                    if (rscSuccess && ugcSuccess) {
                        const message = ugcMessage ? `RSC_Theme文件保存成功，${ugcMessage}` : 'RSC_Theme文件已成功保存到原位置';
                        App.Utils.showStatus(message, 'success');
                        return;
                    } else if (rscSuccess && !ugcSuccess) {
                        // RSC成功但UGC失败，提供重新选择UGC文件的选项
                        const retryMessage = 'RSC_Theme文件保存成功，但UGCTheme文件保存失败。是否重新选择UGCTheme文件？';
                        if (confirm(retryMessage)) {
                            // 重新选择UGC文件
                            const ugcRetrySuccess = await enableUGCDirectFileSave();
                            if (ugcRetrySuccess && ugcResult && ugcResult.workbook) {
                                // 重新尝试保存UGC文件
                                const ugcRetryResult = await saveUGCFileDirectly(ugcResult.workbook);
                                if (ugcRetryResult) {
                                    App.Utils.showStatus('所有文件已成功保存', 'success');
                                } else {
                                    App.Utils.showStatus('UGCTheme文件仍然保存失败，请检查文件状态', 'error');
                                }
                            } else {
                                App.Utils.showStatus('RSC_Theme文件已保存，UGCTheme文件未重新选择', 'warning');
                            }
                        } else {
                            App.Utils.showStatus('RSC_Theme文件已保存，UGCTheme文件保存失败', 'warning');
                        }
                        return;
                    } else {
                        App.Utils.showStatus('文件保存失败，请检查控制台错误信息', 'error');
                        return;
                    }
                }
            }

            // 传统下载方式
            await downloadFileTraditionally(workbook, themeName, ugcResult);
        } catch (error) {
            console.error('文件保存失败:', error);
            App.Utils.showStatus('文件保存失败: ' + error.message, 'error');
        }
    }

    /**
     * 保存UGC文件到原位置
     * @param {Object} workbook - 更新后的UGC工作簿
     */
    async function saveUGCFileDirectly(workbook) {
        try {
            if (!ugcThemeData || !ugcThemeData.fileHandle) {
                console.error('UGC文件句柄不存在');
                App.Utils.showStatus('UGC文件句柄不存在，请重新选择UGCTheme文件', 'error');
                return false;
            }

            console.log('开始保存UGC文件到原位置...');

            // 检查文件句柄权限
            const fileHandle = ugcThemeData.fileHandle;
            try {
                const permission = await fileHandle.queryPermission({ mode: 'readwrite' });
                console.log('UGC文件当前权限:', permission);

                if (permission !== 'granted') {
                    console.log('尝试重新请求UGC文件权限...');
                    const newPermission = await fileHandle.requestPermission({ mode: 'readwrite' });
                    if (newPermission !== 'granted') {
                        console.error('无法获取UGC文件写入权限');
                        App.Utils.showStatus('无法获取UGC文件写入权限，请重新选择文件', 'error');
                        return false;
                    }
                }
            } catch (permissionError) {
                console.error('UGC文件权限检查失败:', permissionError);
                App.Utils.showStatus('UGC文件权限检查失败，文件可能已被移动或删除，请重新选择', 'error');
                return false;
            }

            // 验证文件是否仍然存在
            try {
                const file = await fileHandle.getFile();
                console.log('UGC文件验证成功:', file.name, file.size, 'bytes');
            } catch (fileError) {
                console.error('UGC文件验证失败:', fileError);
                App.Utils.showStatus('UGC文件已不存在或无法访问，请重新选择文件', 'error');
                return false;
            }

            // 生成Excel文件数据 (使用xls格式以兼容Unity工具)
            console.log('UGC原文件名:', ugcThemeData.fileName);
            const cleanWorkbook = createCleanWorkbook(workbook);
            const excelBuffer = XLSX.write(cleanWorkbook, {
                bookType: 'xls',
                type: 'array'
            });
            console.log('UGC文件已强制转换为.xls格式');

            console.log('UGC文件数据大小:', excelBuffer.byteLength, 'bytes');

            // 写入文件
            const writable = await fileHandle.createWritable();
            await writable.write(excelBuffer);
            await writable.close();

            console.log('✅ UGC文件保存成功');

            // UGC文件保存成功后刷新数据预览
            console.log('UGC文件保存成功，开始刷新数据预览...');
            refreshDataPreview();

            return true;
        } catch (error) {
            console.error('UGC文件保存失败:', error);

            // 根据错误类型提供不同的用户提示
            if (error.name === 'InvalidStateError') {
                App.Utils.showStatus('UGC文件状态已变更，请重新选择UGCTheme文件后再试', 'error');
            } else if (error.name === 'NotAllowedError') {
                App.Utils.showStatus('没有UGC文件写入权限，请重新选择文件并授权', 'error');
            } else if (error.name === 'NotFoundError') {
                App.Utils.showStatus('UGC文件未找到，可能已被移动或删除', 'error');
            } else {
                App.Utils.showStatus('UGC文件保存失败: ' + error.message, 'error');
            }

            return false;
        }
    }

    /**
     * 传统下载方式
     * @param {Object} workbook - 更新后的RSC工作簿
     * @param {string} themeName - 主题名称
     * @param {Object} ugcResult - UGC处理结果
     */
    async function downloadFileTraditionally(workbook, themeName, ugcResult) {
        try {
            // 下载RSC_Theme文件 (使用xls格式以兼容Unity工具)
            const cleanWorkbook = createCleanWorkbook(workbook);
            const excelBuffer = XLSX.write(cleanWorkbook, {
                bookType: 'xls',
                type: 'array'
            });

            const blob = new Blob([excelBuffer], {
                type: 'application/vnd.ms-excel'
            });

            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `RSC_Theme_Updated_${themeName}_${new Date().toISOString().slice(0, 10)}.xls`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

            setTimeout(() => {
                URL.revokeObjectURL(url);
            }, 100);

            // 下载UGC文件（如果有处理结果）
            if (ugcResult && ugcResult.success && ugcResult.workbook) {
                console.log('开始下载UGC文件...');

                const cleanUgcWorkbook = createCleanWorkbook(ugcResult.workbook);
                const ugcExcelBuffer = XLSX.write(cleanUgcWorkbook, {
                    bookType: 'xls',
                    type: 'array'
                });

                const ugcBlob = new Blob([ugcExcelBuffer], {
                    type: 'application/vnd.ms-excel'
                });

                const ugcUrl = URL.createObjectURL(ugcBlob);
                const ugcLink = document.createElement('a');
                ugcLink.href = ugcUrl;
                ugcLink.download = `UGCTheme_Updated_${themeName}_${new Date().toISOString().slice(0, 10)}.xls`;
                document.body.appendChild(ugcLink);
                ugcLink.click();
                document.body.removeChild(ugcLink);

                setTimeout(() => {
                    URL.revokeObjectURL(ugcUrl);
                }, 200);

                App.Utils.showStatus('所有文件下载成功', 'success');
            } else {
                App.Utils.showStatus('RSC文件下载成功', 'success');
            }

            // 下载完成后刷新数据预览
            console.log('文件下载完成，开始刷新数据预览...');
            refreshDataPreview();
        } catch (error) {
            console.error('文件下载失败:', error);
            App.Utils.showStatus('文件下载失败: ' + error.message, 'error');
        }
    }



    /**
     * 确认是否直接保存到原文件
     */
    async function confirmDirectSave() {
        return new Promise((resolve) => {
            const modal = document.createElement('div');
            modal.style.cssText = `
                position: fixed; top: 0; left: 0; width: 100%; height: 100%;
                background: rgba(0,0,0,0.5); display: flex; align-items: center;
                justify-content: center; z-index: 10000;
            `;

            modal.innerHTML = `
                <div style="background: white; padding: 30px; border-radius: 10px; max-width: 500px; text-align: center;">
                    <h3>🎯 保存方式选择</h3>
                    <p><strong>您希望如何保存更新后的文件？</strong></p>
                    <div style="margin: 20px 0;">
                        <button id="saveDirectBtn" style="background: #28a745; color: white; border: none; padding: 10px 20px; margin: 5px; border-radius: 5px; cursor: pointer;">
                            ✅ 直接保存到原文件
                        </button>
                        <button id="downloadBtn" style="background: #007bff; color: white; border: none; padding: 10px 20px; margin: 5px; border-radius: 5px; cursor: pointer;">
                            📥 下载新文件
                        </button>
                    </div>
                    <p style="font-size: 12px; color: #666;">
                        直接保存：立即覆盖原文件（推荐）<br>
                        下载新文件：传统方式，需要手动替换
                    </p>
                </div>
            `;

            document.body.appendChild(modal);

            document.getElementById('saveDirectBtn').onclick = () => {
                document.body.removeChild(modal);
                resolve(true);
            };

            document.getElementById('downloadBtn').onclick = () => {
                document.body.removeChild(modal);
                resolve(false);
            };
        });
    }

    /**
     * 创建清理后的工作簿
     * 移除SheetJS无法处理的复杂属性，只保留纯数据
     * @param {Object} originalWorkbook - 原始工作簿
     * @returns {Object} 清理后的工作簿
     */
    function createCleanWorkbook(originalWorkbook) {
        console.log('开始清理工作簿数据...');
        const cleanWb = XLSX.utils.book_new();

        // 清理并重建每个工作表
        originalWorkbook.SheetNames.forEach(sheetName => {
            const worksheet = originalWorkbook.Sheets[sheetName];

            // 提取纯数据（移除所有格式和特殊属性）
            const data = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                defval: '',
                raw: false  // 确保数据为字符串格式
            });

            console.log(`清理工作表 "${sheetName}"，数据行数: ${data.length}`);

            // 从纯数据重建工作表
            const newWorksheet = XLSX.utils.aoa_to_sheet(data);

            // 添加到新工作簿
            XLSX.utils.book_append_sheet(cleanWb, newWorksheet, sheetName);
        });

        console.log('工作簿清理完成，工作表数量:', cleanWb.SheetNames.length);
        return cleanWb;
    }

    /**
     * 直接保存文件到原位置（增强版，包含详细调试）
     */
    async function saveFileDirectly(workbook) {
        console.log('=== 开始直接保存文件到原位置 ===');

        try {
            const fileHandle = rscThemeData.fileHandle;

            if (!fileHandle) {
                throw new Error('文件句柄不存在，请重新选择文件');
            }

            // 保存前的工作簿验证
            console.log('=== 保存前工作簿验证 ===');
            if (!workbook) {
                throw new Error('工作簿对象为空');
            }

            console.log('工作簿验证:');
            console.log('- 工作表数量:', workbook.SheetNames ? workbook.SheetNames.length : 'undefined');
            console.log('- 工作表名称:', workbook.SheetNames);

            if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                throw new Error('工作簿中没有工作表');
            }

            const firstSheetName = workbook.SheetNames[0];
            const firstSheet = workbook.Sheets[firstSheetName];

            if (!firstSheet) {
                throw new Error(`工作表 "${firstSheetName}" 不存在`);
            }

            console.log('- 第一个工作表范围:', firstSheet['!ref']);

            // 验证工作表数据
            if (!firstSheet['!ref']) {
                console.warn('⚠️ 工作表没有范围信息，可能为空');
            } else {
                const range = XLSX.utils.decode_range(firstSheet['!ref']);
                console.log('- 工作表行数:', range.e.r + 1);
                console.log('- 工作表列数:', range.e.c + 1);

                if (range.e.r < 0 || range.e.c < 0) {
                    throw new Error('工作表范围无效');
                }
            }

            console.log('文件句柄信息:', {
                name: fileHandle.name,
                kind: fileHandle.kind
            });

            // 验证权限
            console.log('检查文件权限...');
            const permission = await fileHandle.queryPermission({ mode: 'readwrite' });
            console.log('当前权限状态:', permission);

            if (permission !== 'granted') {
                console.log('请求写入权限...');
                const newPermission = await fileHandle.requestPermission({ mode: 'readwrite' });
                console.log('新权限状态:', newPermission);

                if (newPermission !== 'granted') {
                    throw new Error('无法获取文件写入权限');
                }
            }

            // 清理工作簿数据并生成Excel数据
            console.log('开始清理工作簿数据...');
            const cleanWorkbook = createCleanWorkbook(workbook);

            console.log('生成Excel数据...');
            const excelBuffer = XLSX.write(cleanWorkbook, {
                bookType: 'xls',
                type: 'array'
            });
            console.log('Excel数据大小:', excelBuffer.byteLength, 'bytes');

            // 创建可写流
            console.log('创建可写流...');
            const writable = await fileHandle.createWritable();
            console.log('可写流创建成功');

            // 写入数据
            console.log('开始写入数据...');
            await writable.write(excelBuffer);
            console.log('数据写入完成');

            // 关闭流
            console.log('关闭可写流...');
            await writable.close();
            console.log('可写流关闭完成');

            // 验证保存结果
            console.log('验证保存结果...');
            const verifyResult = await verifySaveResult(fileHandle, workbook);

            if (verifyResult) {
                App.Utils.showStatus('✅ 文件已直接保存到原位置！验证通过。', 'success');
                showDirectSaveSuccess();

                // 保存成功后刷新数据预览
                console.log('保存成功，开始刷新数据预览...');
                refreshDataPreview();
            } else {
                console.warn('⚠️ 保存验证失败，但写入操作已完成');
                App.Utils.showStatus('⚠️ 文件已写入，但验证失败。请手动检查文件内容。', 'warning');
                showSaveVerificationWarning();

                // 即使验证失败，也尝试刷新数据预览
                console.log('保存完成但验证失败，仍尝试刷新数据预览...');
                refreshDataPreview();
            }

            console.log('=== 直接保存文件完成 ===');
            return true;

        } catch (error) {
            console.error('=== 直接保存失败 ===');
            console.error('错误详情:', error);
            console.error('错误堆栈:', error.stack);

            App.Utils.showStatus('直接保存失败，将使用下载方式: ' + error.message, 'warning');

            // 显示详细错误信息
            showSaveErrorDetails(error);

            // 回退到下载方式 (使用xls格式以兼容Unity工具)
            console.log('回退到传统下载方式...');
            const cleanWorkbook = createCleanWorkbook(workbook);
            const wbout = XLSX.write(cleanWorkbook, { bookType: 'xls', type: 'array' });
            const blob = new Blob([wbout], { type: 'application/octet-stream' });

            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = rscThemeData.fileName || 'RSC_Theme_Updated.xls';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);

            setTimeout(() => {
                URL.revokeObjectURL(url);
            }, 100);

            return false;
        }
    }

    /**
     * 验证保存结果
     */
    async function verifySaveResult(fileHandle, originalWorkbook) {
        try {
            console.log('开始验证保存结果...');

            // 重新读取文件
            const file = await fileHandle.getFile();
            console.log('重新读取文件信息:', {
                name: file.name,
                size: file.size,
                lastModified: new Date(file.lastModified).toLocaleString()
            });

            // 解析重新读取的文件
            const arrayBuffer = await file.arrayBuffer();
            const reloadedWorkbook = XLSX.read(arrayBuffer, { type: 'array' });
            const reloadedWorksheet = reloadedWorkbook.Sheets[reloadedWorkbook.SheetNames[0]];
            const reloadedData = XLSX.utils.sheet_to_json(reloadedWorksheet, {
                header: 1,
                defval: '',
                raw: false
            });

            console.log('重新读取的数据行数:', reloadedData.length);
            console.log('原始数据行数:', rscThemeData.data.length);

            // 比较数据
            if (reloadedData.length === rscThemeData.data.length) {
                console.log('✅ 数据行数匹配');

                // 检查最后一行（通常是新添加的主题行）
                const lastRowIndex = reloadedData.length - 1;
                const reloadedLastRow = reloadedData[lastRowIndex];
                const originalLastRow = rscThemeData.data[lastRowIndex];

                console.log('重新读取的最后一行:', reloadedLastRow);
                console.log('原始最后一行:', originalLastRow);

                const rowsMatch = JSON.stringify(reloadedLastRow) === JSON.stringify(originalLastRow);
                if (rowsMatch) {
                    console.log('✅ 数据内容验证通过');
                    return true;
                } else {
                    console.warn('⚠️ 数据内容不匹配');
                    return false;
                }
            } else {
                console.warn('⚠️ 数据行数不匹配');
                return false;
            }

        } catch (error) {
            console.error('验证保存结果失败:', error);
            return false;
        }
    }

    /**
     * 显示保存错误详情
     */
    function showSaveErrorDetails(error) {
        const errorModal = document.createElement('div');
        errorModal.style.cssText = `
            position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(0,0,0,0.5); display: flex; align-items: center;
            justify-content: center; z-index: 10000;
        `;

        errorModal.innerHTML = `
            <div style="background: white; padding: 30px; border-radius: 10px; max-width: 600px; max-height: 80vh; overflow-y: auto;">
                <h3 style="color: #dc3545;">❌ 直接保存失败</h3>
                <p><strong>错误信息：</strong> ${error.message}</p>
                <p><strong>可能的原因：</strong></p>
                <ul>
                    <li>文件被其他程序占用（如Excel正在打开该文件）</li>
                    <li>文件权限不足</li>
                    <li>磁盘空间不足</li>
                    <li>浏览器权限被撤销</li>
                </ul>
                <p><strong>解决建议：</strong></p>
                <ul>
                    <li>确保RSC_Theme.xls文件没有被Excel或其他程序打开</li>
                    <li>检查文件是否为只读状态</li>
                    <li>重新选择文件并授权权限</li>
                    <li>使用下载方式作为备选方案</li>
                </ul>
                <div style="text-align: center; margin-top: 20px;">
                    <button onclick="this.parentElement.parentElement.parentElement.remove()"
                            style="background: #007bff; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer;">
                        确定
                    </button>
                </div>
            </div>
        `;

        document.body.appendChild(errorModal);
    }





    /**
     * 显示保存验证警告
     */
    function showSaveVerificationWarning() {
        const warningModal = document.createElement('div');
        warningModal.style.cssText = `
            position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%);
            background: #fff3cd; border: 2px solid #ffc107; padding: 20px;
            border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            z-index: 10000; max-width: 500px;
        `;

        warningModal.innerHTML = `
            <div style="text-align: center;">
                <h3 style="color: #856404; margin: 0 0 15px 0;">⚠️ 保存验证警告</h3>
                <p style="margin: 10px 0; color: #856404;">
                    文件已写入，但验证过程中发现问题。<br>
                    请手动检查文件内容是否正确更新。
                </p>
                <div style="margin-top: 20px;">
                    <button onclick="this.parentElement.parentElement.parentElement.remove()"
                            style="background: #ffc107; color: #000; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; margin: 0 5px;">
                        我知道了
                    </button>
                </div>
            </div>
        `;

        document.body.appendChild(warningModal);

        // 3秒后自动关闭
        setTimeout(() => {
            if (warningModal.parentNode) {
                warningModal.remove();
            }
        }, 5000);
    }

    /**
     * 显示直接保存成功的提示
     */
    function showDirectSaveSuccess() {
        const successModal = document.createElement('div');
        successModal.style.cssText = `
            position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%);
            background: #d4edda; border: 2px solid #28a745; padding: 20px;
            border-radius: 10px; z-index: 10001; text-align: center;
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        `;

        successModal.innerHTML = `
            <h3 style="color: #155724; margin: 0 0 10px 0;">🎉 保存成功！</h3>
            <p style="color: #155724; margin: 0;">
                主题表已直接更新！<br>
                您可以立即在Unity中新建主题关查看颜色。
            </p>
        `;

        document.body.appendChild(successModal);

        // 3秒后自动关闭
        setTimeout(() => {
            if (document.body.contains(successModal)) {
                document.body.removeChild(successModal);
            }
        }, 3000);
    }

    /**
     * 重置所有状态
     */
    function resetAll() {
        sourceData = null;
        unityProjectFiles = null;
        rscThemeData = null;
        ugcThemeData = null;
        processedResult = null;
        rscAllSheetsData = null;
        ugcAllSheetsData = null;

        // 重置UI状态
        if (themeNameInput) {
            themeNameInput.value = '';
        }

        // 隐藏相关区域
        const sections = ['themeInputSection', 'resultDisplay', 'dataPreview'];
        sections.forEach(sectionId => {
            const section = document.getElementById(sectionId);
            if (section) {
                section.style.display = 'none';
            }
        });

        // 重置文件状态
        updateFileStatus('sourceFileStatus', '未选择', '');
        updateFileStatus('rscThemeStatus', '未找到', '');
        updateFileStatus('ugcThemeStatus', '未找到', '');

        App.Utils.showStatus('已重置所有状态', 'info');
    }

    /**
     * 验证所有颜色通道的处理完整性
     * @param {Array} updatedColors - 更新的颜色数组
     * @param {number} rowIndex - 主题行索引
     */
    function validateColorChannelCompleteness(updatedColors, rowIndex) {
        console.log('=== 开始验证颜色通道处理完整性 ===');

        if (!rscThemeData || !rscThemeData.data || !rscThemeData.data[rowIndex]) {
            console.error('❌ 验证失败：RSC主题数据不完整');
            return false;
        }

        const headerRow = rscThemeData.data[0];
        const themeRow = rscThemeData.data[rowIndex];

        // 获取所有应该处理的颜色通道
        const expectedChannels = mappingData.data
            .filter(mapping => {
                const rcChannel = mapping['RC现在的主题通道'];
                return rcChannel && rcChannel !== '占不导入' && rcChannel !== '' && rcChannel !== '暂不导入';
            })
            .map(mapping => mapping['RC现在的主题通道']);

        console.log('预期处理的颜色通道:', expectedChannels);
        console.log('实际处理的颜色通道:', updatedColors.map(c => c.channel));

        let allChannelsProcessed = true;
        let channelsWithDefaults = 0;
        let channelsWithValues = 0;

        expectedChannels.forEach(channel => {
            const columnIndex = headerRow.findIndex(col => col === channel);
            const updatedColor = updatedColors.find(c => c.channel === channel);

            if (columnIndex === -1) {
                console.error(`❌ 未找到颜色通道列: ${channel}`);
                allChannelsProcessed = false;
                return;
            }

            if (!updatedColor) {
                console.error(`❌ 颜色通道未处理: ${channel}`);
                allChannelsProcessed = false;
                return;
            }

            const actualValue = themeRow[columnIndex];
            if (!actualValue || actualValue === '') {
                console.error(`❌ 颜色通道值为空: ${channel}`);
                allChannelsProcessed = false;
                return;
            }

            if (updatedColor.isDefault) {
                channelsWithDefaults++;
                console.log(`⚠️ 使用默认值: ${channel} = ${actualValue}`);
            } else {
                channelsWithValues++;
                console.log(`✅ 使用源数据值: ${channel} = ${actualValue}`);
            }
        });

        console.log(`处理结果统计: 总计${expectedChannels.length}个通道, 源数据值${channelsWithValues}个, 默认值${channelsWithDefaults}个`);

        if (allChannelsProcessed) {
            console.log('✅ 所有颜色通道处理完整性验证通过');
        } else {
            console.error('❌ 颜色通道处理完整性验证失败');
        }

        console.log('=== 颜色通道处理完整性验证完成 ===');
        return allChannelsProcessed;
    }

    /**
     * 检查和输出当前RSC数据状态（调试用）
     */
    function debugRSCDataState() {
        console.log('=== RSC数据状态检查 ===');

        if (!rscThemeData || !rscThemeData.data) {
            console.error('❌ RSC数据未加载');
            return;
        }

        const data = rscThemeData.data;
        console.log(`总行数: ${data.length}`);

        if (data.length > 0) {
            console.log('表头:', data[0]);

            // 输出所有数据行
            for (let i = 1; i < data.length; i++) {
                console.log(`行${i}:`, data[i]);

                // 检查notes字段
                const notesColumnIndex = data[0].findIndex(col => col === 'notes');
                if (notesColumnIndex !== -1 && data[i][notesColumnIndex]) {
                    console.log(`  主题名称: ${data[i][notesColumnIndex]}`);
                }
            }
        }

        console.log('=== RSC数据状态检查完成 ===');
    }



    /**
     * 获取对比色（用于颜色值显示）
     */
    function getContrastColor(hexColor) {
        const r = parseInt(hexColor.substr(0, 2), 16);
        const g = parseInt(hexColor.substr(2, 2), 16);
        const b = parseInt(hexColor.substr(4, 2), 16);
        const brightness = (r * 299 + g * 587 + b * 114) / 1000;
        return brightness > 128 ? '#000000' : '#FFFFFF';
    }

    /**
     * 启用直接文件保存功能（File System Access API）
     */
    async function enableDirectFileSave() {
        if (!('showOpenFilePicker' in window)) {
            App.Utils.showStatus('当前浏览器不支持直接文件保存，请使用Chrome 86+或Edge 86+', 'warning');
            return false;
        }

        try {
            // 获取记忆的文件信息
            const lastFileInfo = App.Utils.getLastPath('RSC_THEME');

            // 构建文件选择器选项 (只接受.xls格式以确保Unity工具兼容性)
            const pickerOptions = {
                types: [{
                    description: 'Excel files (.xls only for Unity compatibility)',
                    accept: {
                        'application/vnd.ms-excel': ['.xls']
                    }
                }],
                multiple: false,
                startIn: App.Utils.getRecommendedStartIn('RSC_THEME')
            };

            // 如果有上次选择的文件，在状态中显示提示
            if (lastFileInfo && lastFileInfo.fileName) {
                console.log(`上次选择的RSC_Theme文件: ${lastFileInfo.fileName}`);
                App.Utils.showStatus(`提示：上次选择的文件是 ${lastFileInfo.fileName}`, 'info', 2000);
            }

            // 选择RSC_Theme文件并获取写入权限
            const [fileHandle] = await window.showOpenFilePicker(pickerOptions);

            // 验证文件格式
            if (!fileHandle.name.toLowerCase().endsWith('.xls')) {
                App.Utils.showStatus('请选择.xls格式的RSC_Theme文件以确保Unity工具兼容性', 'error');
                return false;
            }

            // 保存路径记忆
            if (fileHandle.name) {
                App.Utils.saveLastPath('RSC_THEME', fileHandle.name);
            }

            // 请求写入权限
            const permission = await fileHandle.requestPermission({ mode: 'readwrite' });
            if (permission !== 'granted') {
                App.Utils.showStatus('无法获取文件写入权限', 'error');
                return false;
            }

            // 读取文件内容
            const file = await fileHandle.getFile();
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                defval: '',
                raw: false
            });

            // 设置RSC数据并保存文件句柄
            rscThemeData = {
                workbook: workbook,
                data: jsonData,
                fileName: file.name,
                fileHandle: fileHandle  // 保存文件句柄用于直接写入
            };

            // 存储所有Sheet数据
            rscAllSheetsData = {};
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const sheetData = XLSX.utils.sheet_to_json(worksheet, {
                    header: 1,
                    defval: '',
                    raw: false
                });
                rscAllSheetsData[sheetName] = sheetData;
            });

            updateFileStatus('rscThemeStatus', `已加载 (支持直接保存): ${file.name}`, 'success');
            App.Utils.showStatus('RSC_Theme文件已加载，支持直接保存到原位置', 'success');
            checkReadyState();

            return true;
        } catch (error) {
            console.error('启用直接文件保存失败:', error);
            App.Utils.showStatus('启用直接文件保存失败: ' + error.message, 'error');
            return false;
        }
    }

    /**
     * 启用UGC直接文件保存功能（File System Access API）
     */
    async function enableUGCDirectFileSave() {
        if (!('showOpenFilePicker' in window)) {
            App.Utils.showStatus('当前浏览器不支持直接文件保存，请使用Chrome 86+或Edge 86+', 'warning');
            return false;
        }

        try {
            // 获取记忆的文件信息
            const lastFileInfo = App.Utils.getLastPath('UGC_THEME');

            // 构建文件选择器选项 (只接受.xls格式以确保Unity工具兼容性)
            const pickerOptions = {
                types: [{
                    description: 'Excel files (.xls only for Unity compatibility)',
                    accept: {
                        'application/vnd.ms-excel': ['.xls']
                    }
                }],
                multiple: false,
                startIn: App.Utils.getRecommendedStartIn('UGC_THEME')
            };

            // 如果有上次选择的文件，在状态中显示提示
            if (lastFileInfo && lastFileInfo.fileName) {
                console.log(`上次选择的UGC_Theme文件: ${lastFileInfo.fileName}`);
                App.Utils.showStatus(`提示：上次选择的文件是 ${lastFileInfo.fileName}`, 'info', 2000);
            }

            // 选择UGCTheme文件并获取写入权限
            const [fileHandle] = await window.showOpenFilePicker(pickerOptions);

            // 验证文件格式
            if (!fileHandle.name.toLowerCase().endsWith('.xls')) {
                App.Utils.showStatus('请选择.xls格式的UGCTheme文件以确保Unity工具兼容性', 'error');
                return false;
            }

            // 保存路径记忆
            if (fileHandle.name) {
                App.Utils.saveLastPath('UGC_THEME', fileHandle.name);
            }

            // 请求写入权限
            const permission = await fileHandle.requestPermission({ mode: 'readwrite' });
            if (permission !== 'granted') {
                App.Utils.showStatus('无法获取文件写入权限', 'error');
                return false;
            }

            // 读取文件内容
            const file = await fileHandle.getFile();
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });

            // 设置UGC数据并保存文件句柄
            ugcThemeData = {
                workbook: workbook,
                fileName: file.name,
                fileHandle: fileHandle  // 保存文件句柄用于直接写入
            };

            // 存储UGC所有Sheet数据
            ugcAllSheetsData = {};
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const sheetData = XLSX.utils.sheet_to_json(worksheet, {
                    header: 1,
                    defval: '',
                    raw: false
                });
                ugcAllSheetsData[sheetName] = sheetData;
            });

            console.log('UGCTheme所有Sheet数据已存储:', Object.keys(ugcAllSheetsData));

            updateFileStatus('ugcThemeStatus', `已加载 (支持直接保存): ${file.name}`, 'success');
            App.Utils.showStatus('UGCTheme文件已加载，支持直接保存到原位置', 'success');
            checkReadyState();

            return true;
        } catch (error) {
            console.error('启用UGC直接文件保存失败:', error);
            App.Utils.showStatus('启用UGC直接文件保存失败: ' + error.message, 'error');
            return false;
        }
    }

    /**
     * 处理文件类型选择变化
     */
    function handleFileTypeChange() {
        console.log('文件类型选择发生变化');
        populateSheetSelector();
    }

    /**
     * 初始化Sheet选择器
     */
    function initializeSheetSelector() {
        console.log('=== 开始初始化Sheet选择器 ===');

        if (!sheetSelectorSection || !rscSheetSelect || !fileTypeSelect) {
            console.warn('Sheet选择器初始化失败：缺少必要DOM元素');
            return;
        }

        // 显示Sheet选择器区域
        sheetSelectorSection.style.display = 'block';

        // 填充Sheet选择器
        populateSheetSelector();
    }

    /**
     * 填充Sheet选择器选项
     */
    function populateSheetSelector() {
        if (!rscSheetSelect || !fileTypeSelect) {
            console.error('Sheet选择器元素未找到');
            return;
        }

        // 获取当前选择的文件类型
        const fileType = fileTypeSelect.value;
        let currentSheetsData = null;
        let fileName = '';

        if (fileType === 'rsc' && rscAllSheetsData) {
            currentSheetsData = rscAllSheetsData;
            fileName = rscThemeData ? rscThemeData.fileName : 'RSC_Theme';
        } else if (fileType === 'ugc' && ugcAllSheetsData) {
            currentSheetsData = ugcAllSheetsData;
            fileName = ugcThemeData ? ugcThemeData.fileName : 'UGCTheme';
        }

        // 清空现有选项
        rscSheetSelect.innerHTML = '<option value="">请选择工作表</option>';

        if (!currentSheetsData) {
            console.warn(`${fileType === 'rsc' ? 'RSC_Theme' : 'UGCTheme'}数据未加载`);
            hideSheetData();
            if (rscSheetInfo) {
                rscSheetInfo.textContent = `${fileType === 'rsc' ? 'RSC_Theme' : 'UGCTheme'}文件未加载`;
            }
            return;
        }

        // 获取所有Sheet名称
        const sheetNames = Object.keys(currentSheetsData);

        // 添加Sheet选项
        sheetNames.forEach(sheetName => {
            const option = document.createElement('option');
            option.value = sheetName;
            option.textContent = sheetName;
            rscSheetSelect.appendChild(option);
        });

        // 显示Sheet选择器区域
        sheetSelectorSection.style.display = 'block';

        // 默认选择Color Sheet（如果存在）
        if (sheetNames.includes('Color')) {
            rscSheetSelect.value = 'Color';
            displaySelectedSheet('Color');
        } else if (sheetNames.length > 0) {
            // 如果没有Color Sheet，选择第一个
            rscSheetSelect.value = sheetNames[0];
            displaySelectedSheet(sheetNames[0]);
        }

        // 更新Sheet信息
        if (rscSheetInfo) {
            rscSheetInfo.textContent = `共 ${sheetNames.length} 个工作表`;
        }
    }

    /**
     * 处理Sheet选择变化
     */
    function handleSheetSelectionChange() {
        const selectedSheet = rscSheetSelect.value;
        if (selectedSheet) {
            displaySelectedSheet(selectedSheet);
        } else {
            hideSheetData();
        }
    }

    /**
     * 显示选中的Sheet数据
     * @param {string} sheetName - Sheet名称
     */
    function displaySelectedSheet(sheetName) {
        if (!fileTypeSelect) {
            console.error('文件类型选择器未找到');
            return;
        }

        // 获取当前选择的文件类型
        const fileType = fileTypeSelect.value;
        let currentSheetsData = null;

        if (fileType === 'rsc') {
            currentSheetsData = rscAllSheetsData;
        } else if (fileType === 'ugc') {
            currentSheetsData = ugcAllSheetsData;
        }

        if (!currentSheetsData || !currentSheetsData[sheetName]) {
            console.error(`Sheet "${sheetName}" 在${fileType === 'rsc' ? 'RSC_Theme' : 'UGCTheme'}中不存在`);
            hideSheetData();
            return;
        }

        const sheetData = currentSheetsData[sheetName];

        if (!sheetData || sheetData.length === 0) {
            console.warn(`Sheet "${sheetName}" 为空`);
            hideSheetData();
            return;
        }

        // 渲染表格
        renderSheetTable(sheetData);

        // 显示数据统计
        displaySheetStats(sheetData, sheetName);

        // 显示数据容器
        if (sheetDataContainer) {
            sheetDataContainer.style.display = 'block';
        }
    }

    /**
     * 渲染Sheet表格
     * @param {Array} data - Sheet数据
     */
    function renderSheetTable(data) {
        if (!rscSheetTableHead || !rscSheetTableBody) {
            console.error('表格元素未找到');
            return;
        }

        // 清空现有内容
        rscSheetTableHead.innerHTML = '';
        rscSheetTableBody.innerHTML = '';

        if (data.length === 0) {
            return;
        }

        // 创建表头
        const headerRow = document.createElement('tr');
        const headers = data[0] || [];

        // 添加行号列
        const rowNumHeader = document.createElement('th');
        rowNumHeader.textContent = '行号';
        rowNumHeader.style.cssText = 'border: 1px solid #ddd; padding: 5px; background-color: #f5f5f5; min-width: 50px;';
        headerRow.appendChild(rowNumHeader);

        // 添加数据列
        headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header || '';
            th.style.cssText = 'border: 1px solid #ddd; padding: 5px; background-color: #f5f5f5; min-width: 60px;';
            headerRow.appendChild(th);
        });

        rscSheetTableHead.appendChild(headerRow);

        // 创建数据行（跳过表头，显示所有数据）
        const maxRows = data.length - 1;
        for (let i = 1; i <= maxRows; i++) {
            const row = data[i] || [];
            const tr = document.createElement('tr');

            // 添加行号
            const rowNumCell = document.createElement('td');
            rowNumCell.textContent = i;
            rowNumCell.style.cssText = 'border: 1px solid #ddd; padding: 5px; background-color: #f9f9f9; text-align: center; font-weight: bold;';
            tr.appendChild(rowNumCell);

            // 添加数据单元格
            const maxCols = Math.max(headers.length, row.length);
            for (let j = 0; j < maxCols; j++) {
                const td = document.createElement('td');
                const cellValue = row[j] || '';
                td.textContent = cellValue;

                let cellStyle = 'border: 1px solid #ddd; padding: 5px;';

                // 高亮颜色值
                if (cellValue && /^[0-9A-F]{6}$/i.test(cellValue)) {
                    cellStyle += ` background-color: #${cellValue}; color: ${getContrastColor(cellValue)};`;
                }

                // 高亮notes列
                if (headers[j] === 'notes' && cellValue) {
                    cellStyle += ' background-color: #e8f5e8; font-weight: bold;';
                }

                td.style.cssText = cellStyle;
                tr.appendChild(td);
            }

            rscSheetTableBody.appendChild(tr);
        }


    }

    /**
     * 显示Sheet统计信息
     * @param {Array} data - Sheet数据
     * @param {string} sheetName - Sheet名称
     */
    function displaySheetStats(data, sheetName) {
        if (!sheetDataStats) {
            return;
        }

        const totalRows = data.length - 1; // 减去表头行
        const totalColumns = data.length > 0 ? Math.max(...data.map(row => row.length)) : 0;
        const displayedRows = Math.min(totalRows, 100);

        sheetDataStats.innerHTML = `
            <strong>工作表：${sheetName}</strong> |
            总行数: ${totalRows} |
            总列数: ${totalColumns} |
            显示行数: ${displayedRows}
        `;
    }

    /**
     * 隐藏Sheet数据
     */
    function hideSheetData() {
        if (sheetDataContainer) {
            sheetDataContainer.style.display = 'none';
        }
        if (sheetDataStats) {
            sheetDataStats.innerHTML = '';
        }
    }

    /**
     * 同步内存数据状态
     * 在主题处理完成后，更新内存中的数据以保持与工作簿一致
     * @param {Object} workbook - 更新后的RSC工作簿
     * @param {number} themeRowIndex - 主题行索引
     */
    function syncMemoryDataState(workbook, themeRowIndex) {
        try {
            console.log('开始同步RSC内存数据状态...');
            console.log('工作簿对象:', workbook);
            console.log('主题行索引:', themeRowIndex);

            if (!workbook || !rscThemeData) {
                console.warn('工作簿或RSC数据不存在，跳过同步');
                return;
            }

            // 获取主工作表名称（通常是第一个）
            const mainSheetName = workbook.SheetNames[0];
            console.log('主工作表名称:', mainSheetName);

            if (!mainSheetName || !workbook.Sheets[mainSheetName]) {
                console.warn('主工作表不存在，跳过同步');
                return;
            }

            // 从工作簿中读取最新数据
            const updatedSheetData = XLSX.utils.sheet_to_json(workbook.Sheets[mainSheetName], {
                header: 1,
                defval: '',
                raw: false
            });

            console.log('从工作簿读取的最新数据行数:', updatedSheetData.length);

            // 更新rscThemeData.data
            rscThemeData.data = updatedSheetData;
            console.log('已更新rscThemeData.data，新行数:', rscThemeData.data.length);

            // 更新rscAllSheetsData中对应的Sheet数据
            if (rscAllSheetsData && rscAllSheetsData[mainSheetName]) {
                rscAllSheetsData[mainSheetName] = updatedSheetData;
                console.log(`已更新rscAllSheetsData["${mainSheetName}"]`);
            }

            // 同步其他工作表（如果有更新）
            workbook.SheetNames.forEach(sheetName => {
                if (sheetName !== mainSheetName && workbook.Sheets[sheetName]) {
                    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
                        header: 1,
                        defval: '',
                        raw: false
                    });

                    if (rscAllSheetsData) {
                        rscAllSheetsData[sheetName] = sheetData;
                        console.log(`已同步工作表 "${sheetName}" 数据`);
                    }
                }
            });

            console.log('✅ RSC内存数据状态同步完成');

        } catch (error) {
            console.error('RSC内存数据状态同步失败:', error);
        }
    }

    /**
     * 同步UGC内存数据状态
     * 在UGC主题处理完成后，更新内存中的UGC数据以保持与工作簿一致
     * @param {Object} workbook - 更新后的UGC工作簿
     */
    function syncUGCMemoryDataState(workbook) {
        try {
            console.log('开始同步UGC内存数据状态...');
            console.log('UGC工作簿对象:', workbook);

            if (!workbook || !ugcThemeData) {
                console.warn('UGC工作簿或UGC数据不存在，跳过同步');
                return;
            }

            // 更新ugcThemeData.workbook
            ugcThemeData.workbook = workbook;
            console.log('已更新ugcThemeData.workbook');

            // 重新生成ugcAllSheetsData
            ugcAllSheetsData = {};
            workbook.SheetNames.forEach(sheetName => {
                if (workbook.Sheets[sheetName]) {
                    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
                        header: 1,
                        defval: '',
                        raw: false
                    });

                    ugcAllSheetsData[sheetName] = sheetData;
                    console.log(`已同步UGC工作表 "${sheetName}" 数据，行数: ${sheetData.length}`);
                }
            });

            console.log('✅ UGC内存数据状态同步完成');
            console.log('UGC工作表列表:', Object.keys(ugcAllSheetsData));

        } catch (error) {
            console.error('UGC内存数据状态同步失败:', error);
        }
    }

    /**
     * 刷新数据预览区域
     * 强制刷新当前显示的数据预览
     */
    function refreshDataPreview() {
        try {
            console.log('开始刷新数据预览区域...');

            // 检查是否有Sheet选择器显示
            const sheetSelectorSection = document.getElementById('sheetSelectorSection');
            if (sheetSelectorSection && sheetSelectorSection.style.display !== 'none') {
                // 获取当前选中的文件类型和Sheet
                const fileTypeSelect = document.getElementById('fileTypeSelect');
                const rscSheetSelect = document.getElementById('rscSheetSelect');
                const currentFileType = fileTypeSelect ? fileTypeSelect.value : 'rsc';
                const currentSheet = rscSheetSelect ? rscSheetSelect.value : null;

                console.log('当前文件类型:', currentFileType);
                console.log('当前选中的Sheet:', currentSheet);

                // 根据文件类型检查数据是否存在
                let hasData = false;
                if (currentFileType === 'rsc' && rscAllSheetsData) {
                    hasData = true;
                } else if (currentFileType === 'ugc' && ugcAllSheetsData) {
                    hasData = true;
                }

                if (hasData) {
                    if (currentSheet) {
                        console.log(`刷新当前选中的${currentFileType.toUpperCase()}Sheet:`, currentSheet);
                        // 重新显示选中的Sheet数据
                        displaySelectedSheet(currentSheet);
                    } else {
                        console.log(`重新初始化${currentFileType.toUpperCase()}Sheet选择器`);
                        // 重新初始化Sheet选择器
                        populateSheetSelector();
                    }
                } else {
                    console.warn(`${currentFileType.toUpperCase()}数据不存在，无法刷新预览`);
                }
            }

            console.log('✅ 数据预览区域刷新完成');

        } catch (error) {
            console.error('数据预览区域刷新失败:', error);
        }
    }

    // 暴露公共接口
    return {
        init: init,
        setSourceData: setSourceData,
        setUnityProjectFiles: setUnityProjectFiles,
        setRSCThemeFile: setRSCThemeFile,
        resetAll: resetAll,
        debugRSCDataState: debugRSCDataState,

        enableDirectFileSave: enableDirectFileSave,
        displayBrowserCompatibility: displayBrowserCompatibility,
        isReady: () => isInitialized,

        // 数据同步功能
        refreshDataPreview: refreshDataPreview
    };



    /**
     * 显示浏览器兼容性信息
     */
    function displayBrowserCompatibility() {
        const compatibilityDiv = document.getElementById('browserCompatibility');
        if (!compatibilityDiv) return;

        const hasFileSystemAccess = 'showOpenFilePicker' in window;
        const userAgent = navigator.userAgent;
        const isChrome = userAgent.includes('Chrome') && !userAgent.includes('Edg');
        const isEdge = userAgent.includes('Edg');
        const isFirefox = userAgent.includes('Firefox');
        const isSafari = userAgent.includes('Safari') && !userAgent.includes('Chrome');

        let html = '';
        let bgColor = '';
        let textColor = '';

        if (hasFileSystemAccess) {
            bgColor = '#d4edda';
            textColor = '#155724';
            html = `
                <strong>🎉 您的浏览器支持直接文件保存功能！</strong><br>
            `;
        } else {
            bgColor = '#fff3cd';
            textColor = '#856404';

            if (isFirefox || isSafari) {
                html = `
                    <strong>⚠️ 当前浏览器不支持直接文件保存</strong><br>
                `;
            } else {
                html = `
                    <strong>⚠️ 浏览器版本可能过旧</strong><br>
                    <small>
                        请更新到 <strong>Chrome 86+</strong> 或 <strong>Edge 86+</strong> 以支持直接文件保存<br>
                        当前将使用传统下载方式，需要手动替换文件
                    </small>
                `;
            }
        }

        compatibilityDiv.style.backgroundColor = bgColor;
        compatibilityDiv.style.color = textColor;
        compatibilityDiv.style.border = `1px solid ${textColor}`;
        compatibilityDiv.innerHTML = html;
    }

})();

// 模块加载完成日志
console.log('ThemeManager模块已加载 - 颜色主题管理功能已就绪');
