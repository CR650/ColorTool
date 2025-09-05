/**
 * 数据导出工具模块
 * 文件说明：负责数据导出功能，支持JSON和CSV格式
 * 创建时间：2025-01-09
 * 依赖：App.Utils, App.DataParser, App.TableRenderer
 */

// 确保全局命名空间存在
window.App = window.App || {};

/**
 * 数据导出工具模块
 * 提供多种格式的数据导出功能
 */
window.App.ExportUtils = (function() {
    'use strict';

    // 模块状态
    let isInitialized = false;

    /**
     * 初始化导出工具模块
     */
    function init() {
        if (isInitialized) {
            console.warn('ExportUtils模块已经初始化');
            return;
        }

        // 检查浏览器是否支持Blob和URL API
        if (!window.Blob || !window.URL) {
            console.error('浏览器不支持文件下载功能');
            App.Utils.showStatus('浏览器不支持文件下载功能', 'error');
            return;
        }

        isInitialized = true;
        console.log('ExportUtils模块初始化完成');
    }

    /**
     * 导出数据
     * 根据指定格式导出当前选中工作表的数据
     * @param {string} format - 导出格式：'json' 或 'csv'
     */
    function exportData(format) {
        if (!App.TableRenderer || !App.DataParser) {
            console.error('依赖模块未加载');
            App.Utils.showStatus('导出功能依赖模块未加载', 'error');
            return;
        }

        const selectedSheet = App.TableRenderer.getCurrentSheet();
        
        if (!selectedSheet) {
            App.Utils.showStatus('请先选择要导出的工作表', 'error');
            return;
        }

        const data = App.DataParser.getSheetData(selectedSheet);
        
        if (!data || data.length === 0) {
            App.Utils.showStatus('没有数据可以导出', 'error');
            return;
        }

        try {
            switch (format.toLowerCase()) {
                case 'json':
                    exportAsJSON(data, selectedSheet);
                    break;
                case 'csv':
                    exportAsCSV(data, selectedSheet);
                    break;
                default:
                    App.Utils.showStatus('不支持的导出格式', 'error');
                    return;
            }
            
            App.Utils.showStatus(`数据已导出为 ${format.toUpperCase()} 格式`, 'success');
        } catch (error) {
            console.error('导出错误:', error);
            App.Utils.showStatus('导出失败：' + error.message, 'error');
        }
    }

    /**
     * 导出为JSON格式
     * 将表格数据转换为JSON对象数组并下载
     * @param {Array} data - 表格数据（二维数组）
     * @param {string} sheetName - 工作表名称
     */
    function exportAsJSON(data, sheetName) {
        if (data.length === 0) {
            throw new Error('数据为空');
        }

        // 获取表头
        const headers = data[0] || [];
        
        // 将数据行转换为对象数组
        const jsonData = data.slice(1).map((row, index) => {
            const obj = {};
            headers.forEach((header, colIndex) => {
                const key = header || `列${colIndex + 1}`;
                obj[key] = row[colIndex] || '';
            });
            return obj;
        });

        // 创建完整的导出对象
        const exportObject = {
            sheetName: sheetName,
            exportTime: new Date().toISOString(),
            totalRows: jsonData.length,
            totalColumns: headers.length,
            headers: headers,
            data: jsonData
        };

        const jsonString = JSON.stringify(exportObject, null, 2);
        const filename = `${sanitizeFilename(sheetName)}.json`;
        
        downloadFile(jsonString, filename, 'application/json');
    }

    /**
     * 导出为CSV格式
     * 将表格数据转换为CSV格式并下载
     * @param {Array} data - 表格数据（二维数组）
     * @param {string} sheetName - 工作表名称
     */
    function exportAsCSV(data, sheetName) {
        if (data.length === 0) {
            throw new Error('数据为空');
        }

        const csvContent = data.map(row => {
            return row.map(cell => formatCSVCell(cell)).join(',');
        }).join('\n');

        const filename = `${sanitizeFilename(sheetName)}.csv`;
        
        // 添加BOM以支持中文字符
        const bom = '\uFEFF';
        const csvWithBom = bom + csvContent;
        
        downloadFile(csvWithBom, filename, 'text/csv;charset=utf-8');
    }

    /**
     * 格式化CSV单元格
     * 处理包含逗号、引号、换行符的单元格内容
     * @param {any} cell - 单元格内容
     * @returns {string} 格式化后的CSV单元格
     */
    function formatCSVCell(cell) {
        const cellStr = String(cell || '');
        
        // 如果包含逗号、引号或换行符，需要用引号包围
        if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n') || cellStr.includes('\r')) {
            // 转义引号（双引号变成两个双引号）
            const escapedCell = cellStr.replace(/"/g, '""');
            return `"${escapedCell}"`;
        }
        
        return cellStr;
    }

    /**
     * 清理文件名
     * 移除或替换文件名中的非法字符
     * @param {string} filename - 原始文件名
     * @returns {string} 清理后的文件名
     */
    function sanitizeFilename(filename) {
        // 移除或替换非法字符
        return filename
            .replace(/[<>:"/\\|?*]/g, '_')  // 替换非法字符为下划线
            .replace(/\s+/g, '_')          // 替换空格为下划线
            .substring(0, 100);            // 限制长度
    }

    /**
     * 下载文件
     * 创建Blob对象并触发下载
     * @param {string} content - 文件内容
     * @param {string} filename - 文件名
     * @param {string} mimeType - MIME类型
     */
    function downloadFile(content, filename, mimeType) {
        try {
            const blob = new Blob([content], { type: mimeType });
            const url = URL.createObjectURL(blob);
            
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            a.style.display = 'none';
            
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            
            // 清理URL对象
            setTimeout(() => {
                URL.revokeObjectURL(url);
            }, 100);
            
        } catch (error) {
            throw new Error('文件下载失败：' + error.message);
        }
    }

    /**
     * 导出所有工作表
     * 将所有工作表数据导出为一个JSON文件
     * @param {string} format - 导出格式，目前只支持'json'
     */
    function exportAllSheets(format = 'json') {
        if (!App.DataParser) {
            App.Utils.showStatus('数据解析模块未加载', 'error');
            return;
        }

        const allData = App.DataParser.getAllData();
        const workbook = App.DataParser.getCurrentWorkbook();
        
        if (!allData || !workbook) {
            App.Utils.showStatus('没有数据可以导出', 'error');
            return;
        }

        try {
            if (format === 'json') {
                exportAllSheetsAsJSON(allData, workbook);
            } else {
                App.Utils.showStatus('批量导出目前只支持JSON格式', 'error');
            }
        } catch (error) {
            console.error('批量导出错误:', error);
            App.Utils.showStatus('批量导出失败：' + error.message, 'error');
        }
    }

    /**
     * 导出所有工作表为JSON格式
     * @param {Object} allData - 所有工作表数据
     * @param {Object} workbook - 工作簿对象
     */
    function exportAllSheetsAsJSON(allData, workbook) {
        const exportObject = {
            exportTime: new Date().toISOString(),
            totalSheets: workbook.SheetNames.length,
            sheets: {}
        };

        workbook.SheetNames.forEach(sheetName => {
            const data = allData[sheetName];
            if (data && data.length > 0) {
                const headers = data[0] || [];
                const jsonData = data.slice(1).map(row => {
                    const obj = {};
                    headers.forEach((header, colIndex) => {
                        const key = header || `列${colIndex + 1}`;
                        obj[key] = row[colIndex] || '';
                    });
                    return obj;
                });

                exportObject.sheets[sheetName] = {
                    headers: headers,
                    totalRows: jsonData.length,
                    totalColumns: headers.length,
                    data: jsonData
                };
            }
        });

        const jsonString = JSON.stringify(exportObject, null, 2);
        const filename = `所有工作表_${new Date().toISOString().split('T')[0]}.json`;
        
        downloadFile(jsonString, filename, 'application/json');
        App.Utils.showStatus('所有工作表已导出为JSON格式', 'success');
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
        exportData: exportData,
        exportAllSheets: exportAllSheets,
        isReady: isReady
    };

})();

// 模块加载完成日志
console.log('ExportUtils模块已加载 - 数据导出功能已就绪');
