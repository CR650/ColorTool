/**
 * 数据解析模块
 * 文件说明：负责Excel和CSV文件的数据解析功能
 * 创建时间：2025-01-09
 * 依赖：App.Utils, App.TableRenderer, XLSX库
 */

// 确保全局命名空间存在
window.App = window.App || {};

/**
 * 数据解析模块
 * 处理Excel和CSV文件的解析，将数据转换为可用格式
 */
window.App.DataParser = (function() {
    'use strict';

    // 模块状态
    let currentWorkbook = null;
    let currentData = null;
    let isInitialized = false;

    /**
     * 初始化数据解析模块
     * 检查依赖库是否可用
     */
    function init() {
        if (isInitialized) {
            console.warn('DataParser模块已经初始化');
            return;
        }

        // 检查XLSX库是否可用
        if (!window.XLSX) {
            console.error('XLSX库未加载，Excel解析功能不可用');
            App.Utils.showStatus('Excel解析库加载失败，请检查网络连接', 'error');
            return;
        }

        isInitialized = true;
        console.log('DataParser模块初始化完成');
    }

    /**
     * 解析文件数据
     * 根据文件类型选择合适的解析方法
     * @param {ArrayBuffer|string} data - 文件数据
     * @param {File} file - 文件对象
     */
    function parseFile(data, file) {
        try {
            if (file.name.toLowerCase().endsWith('.csv')) {
                parseCSVData(data);
            } else {
                parseExcelData(data);
            }
        } catch (error) {
            console.error('数据解析错误:', error);
            App.Utils.showStatus('数据解析失败：' + error.message, 'error');
        }
    }

    /**
     * 解析CSV数据
     * 将CSV文本转换为二维数组格式
     * @param {string} csvText - CSV文件内容
     */
    function parseCSVData(csvText) {
        const lines = csvText.split('\n').filter(line => line.trim());
        
        if (lines.length === 0) {
            App.Utils.showStatus('CSV文件为空', 'error');
            return;
        }

        const data = lines.map(line => parseCSVLine(line));
        
        // 存储解析结果
        currentData = { 'Sheet1': data };
        currentWorkbook = { SheetNames: ['Sheet1'] };
        
        // 更新工作表数量显示
        updateSheetCount(1);
        
        // 通知表格渲染模块
        if (App.TableRenderer) {
            App.TableRenderer.populateSheetSelector(['Sheet1']);
        }
        
        App.Utils.showStatus('CSV文件解析成功', 'success');
    }

    /**
     * 解析CSV行数据
     * 处理包含逗号和引号的CSV格式
     * @param {string} line - CSV行文本
     * @returns {Array} 解析后的单元格数组
     */
    function parseCSVLine(line) {
        const cells = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            const nextChar = line[i + 1];
            
            if (char === '"') {
                if (inQuotes && nextChar === '"') {
                    // 转义的引号
                    current += '"';
                    i++; // 跳过下一个引号
                } else {
                    // 切换引号状态
                    inQuotes = !inQuotes;
                }
            } else if (char === ',' && !inQuotes) {
                // 字段分隔符
                cells.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        
        // 添加最后一个字段
        cells.push(current.trim());
        
        return cells;
    }

    /**
     * 解析Excel数据
     * 使用XLSX库解析Excel文件
     * @param {ArrayBuffer} arrayBuffer - Excel文件的二进制数据
     */
    function parseExcelData(arrayBuffer) {
        if (!window.XLSX) {
            throw new Error('XLSX库未加载');
        }

        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        currentWorkbook = workbook;
        currentData = {};

        // 解析所有工作表
        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1, 
                defval: '',
                raw: false 
            });
            currentData[sheetName] = jsonData;
        });

        // 更新工作表数量显示
        updateSheetCount(workbook.SheetNames.length);
        
        // 通知表格渲染模块
        if (App.TableRenderer) {
            App.TableRenderer.populateSheetSelector(workbook.SheetNames);
        }
        
        App.Utils.showStatus(`Excel文件解析成功，包含 ${workbook.SheetNames.length} 个工作表`, 'success');
    }

    /**
     * 更新工作表数量显示
     * @param {number} count - 工作表数量
     */
    function updateSheetCount(count) {
        const sheetCountElement = document.getElementById('sheetCount');
        if (sheetCountElement) {
            sheetCountElement.textContent = count;
        }
    }

    /**
     * 获取指定工作表的数据
     * @param {string} sheetName - 工作表名称
     * @returns {Array|null} 工作表数据或null
     */
    function getSheetData(sheetName) {
        if (!currentData || !currentData[sheetName]) {
            return null;
        }
        return currentData[sheetName];
    }

    /**
     * 获取所有工作表名称
     * @returns {Array} 工作表名称数组
     */
    function getSheetNames() {
        if (!currentWorkbook || !currentWorkbook.SheetNames) {
            return [];
        }
        return currentWorkbook.SheetNames;
    }

    /**
     * 获取当前解析的所有数据
     * @returns {Object|null} 包含所有工作表数据的对象
     */
    function getAllData() {
        return currentData;
    }

    /**
     * 获取当前工作簿对象
     * @returns {Object|null} 工作簿对象
     */
    function getCurrentWorkbook() {
        return currentWorkbook;
    }

    /**
     * 清除当前数据
     * 重置模块状态
     */
    function clearData() {
        currentWorkbook = null;
        currentData = null;
    }

    /**
     * 验证数据完整性
     * 检查解析后的数据是否有效
     * @param {string} sheetName - 工作表名称
     * @returns {Object} 验证结果
     */
    function validateData(sheetName) {
        const data = getSheetData(sheetName);
        
        if (!data) {
            return {
                isValid: false,
                message: '工作表不存在或数据为空'
            };
        }

        if (data.length === 0) {
            return {
                isValid: false,
                message: '工作表为空'
            };
        }

        const maxColumns = Math.max(...data.map(row => row.length));
        const totalRows = data.length;
        const hasHeader = totalRows > 0;

        return {
            isValid: true,
            totalRows: totalRows,
            totalColumns: maxColumns,
            hasHeader: hasHeader,
            message: '数据验证通过'
        };
    }

    /**
     * 获取模块状态
     * @returns {boolean} 模块是否已初始化
     */
    function isReady() {
        return isInitialized;
    }

    // 暴露公共接口
    return {
        init: init,
        parseFile: parseFile,
        getSheetData: getSheetData,
        getSheetNames: getSheetNames,
        getAllData: getAllData,
        getCurrentWorkbook: getCurrentWorkbook,
        clearData: clearData,
        validateData: validateData,
        isReady: isReady
    };

})();

// 模块加载完成日志
console.log('DataParser模块已加载 - 数据解析功能已就绪');
