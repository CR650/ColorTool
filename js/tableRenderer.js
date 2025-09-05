/**
 * 表格渲染模块
 * 文件说明：负责表格渲染、数据展示、工作表选择等功能
 * 创建时间：2025-01-09
 * 依赖：App.Utils, App.DataParser
 */

// 确保全局命名空间存在
window.App = window.App || {};

/**
 * 表格渲染模块
 * 处理数据表格的渲染和展示功能
 */
window.App.TableRenderer = (function() {
    'use strict';

    // DOM元素引用
    let sheetSelector = null;
    let sheetSelect = null;
    let dataDisplay = null;
    let tableHead = null;
    let tableBody = null;
    let dataStats = null;

    // 模块状态
    let isInitialized = false;
    let maxDisplayRows = 1000; // 最大显示行数，用于性能优化

    /**
     * 初始化表格渲染模块
     * 获取DOM元素引用并设置事件监听器
     */
    function init() {
        if (isInitialized) {
            console.warn('TableRenderer模块已经初始化');
            return;
        }

        // 获取DOM元素
        sheetSelector = document.getElementById('sheetSelector');
        sheetSelect = document.getElementById('sheetSelect');
        dataDisplay = document.getElementById('dataDisplay');
        tableHead = document.getElementById('tableHead');
        tableBody = document.getElementById('tableBody');
        dataStats = document.getElementById('dataStats');

        // 验证必要元素是否存在
        if (!sheetSelector || !sheetSelect || !dataDisplay || !tableHead || !tableBody || !dataStats) {
            console.error('TableRenderer初始化失败：缺少必要的DOM元素');
            return;
        }

        // 设置事件监听器
        setupEventListeners();
        
        isInitialized = true;
        console.log('TableRenderer模块初始化完成');
    }

    /**
     * 设置事件监听器
     */
    function setupEventListeners() {
        // 工作表选择变化事件
        sheetSelect.addEventListener('change', handleSheetChange);
    }

    /**
     * 处理工作表选择变化
     * 当用户选择不同工作表时触发
     */
    function handleSheetChange() {
        const selectedSheet = sheetSelect.value;
        if (selectedSheet) {
            displaySheet(selectedSheet);
        } else {
            hideDataDisplay();
        }
    }

    /**
     * 填充工作表选择器
     * 根据解析的工作表名称填充下拉选择器
     * @param {Array} sheetNames - 工作表名称数组
     */
    function populateSheetSelector(sheetNames) {
        if (!sheetSelect) {
            console.error('工作表选择器元素未找到');
            return;
        }

        // 清空现有选项
        sheetSelect.innerHTML = '<option value="">请选择工作表</option>';
        
        // 添加工作表选项
        sheetNames.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            sheetSelect.appendChild(option);
        });

        // 显示工作表选择器
        sheetSelector.style.display = 'block';
        
        // 如果只有一个工作表，自动选择并显示
        if (sheetNames.length === 1) {
            sheetSelect.value = sheetNames[0];
            displaySheet(sheetNames[0]);
        }
    }

    /**
     * 显示指定工作表的数据
     * @param {string} sheetName - 工作表名称
     */
    function displaySheet(sheetName) {
        if (!App.DataParser) {
            console.error('DataParser模块未加载');
            return;
        }

        const data = App.DataParser.getSheetData(sheetName);
        
        if (!data) {
            App.Utils.showStatus('工作表数据获取失败', 'error');
            hideDataDisplay();
            return;
        }

        if (data.length === 0) {
            App.Utils.showStatus('选中的工作表为空', 'error');
            hideDataDisplay();
            return;
        }

        // 渲染表格
        renderTable(data);
        
        // 显示数据统计
        displayDataStats(data);
        
        // 显示数据展示区域
        dataDisplay.style.display = 'block';
        
        App.Utils.showStatus(`工作表 "${sheetName}" 加载成功`, 'success');
    }

    /**
     * 渲染数据表格
     * @param {Array} data - 二维数组格式的表格数据
     */
    function renderTable(data) {
        if (!tableHead || !tableBody) {
            console.error('表格元素未找到');
            return;
        }

        // 清空现有内容
        tableHead.innerHTML = '';
        tableBody.innerHTML = '';

        if (data.length === 0) {
            return;
        }

        // 计算最大列数
        const maxColumns = Math.max(...data.map(row => row.length));
        
        // 创建表头
        createTableHeader(data[0], maxColumns);
        
        // 创建表格内容（跳过第一行作为表头）
        createTableBody(data.slice(1), maxColumns);
    }

    /**
     * 创建表格表头
     * @param {Array} headerRow - 表头行数据
     * @param {number} maxColumns - 最大列数
     */
    function createTableHeader(headerRow, maxColumns) {
        const headerRowElement = document.createElement('tr');
        
        for (let i = 0; i < maxColumns; i++) {
            const th = document.createElement('th');
            th.textContent = headerRow[i] || `列 ${i + 1}`;
            th.title = headerRow[i] || `列 ${i + 1}`; // 添加提示文本
            headerRowElement.appendChild(th);
        }
        
        tableHead.appendChild(headerRowElement);
    }

    /**
     * 创建表格主体内容
     * @param {Array} bodyData - 表格主体数据
     * @param {number} maxColumns - 最大列数
     */
    function createTableBody(bodyData, maxColumns) {
        const displayRows = Math.min(bodyData.length, maxDisplayRows);
        
        // 创建数据行
        for (let i = 0; i < displayRows; i++) {
            const row = bodyData[i] || [];
            const tr = document.createElement('tr');
            
            for (let j = 0; j < maxColumns; j++) {
                const td = document.createElement('td');
                const cellValue = row[j] || '';
                td.textContent = cellValue;
                td.title = cellValue; // 添加提示文本，便于查看长内容
                tr.appendChild(td);
            }
            
            tableBody.appendChild(tr);
        }

        // 如果数据被截断，显示提示行
        if (bodyData.length > maxDisplayRows) {
            createTruncationNotice(bodyData.length, maxColumns);
        }
    }

    /**
     * 创建数据截断提示
     * @param {number} totalRows - 总行数
     * @param {number} maxColumns - 最大列数
     */
    function createTruncationNotice(totalRows, maxColumns) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        
        td.colSpan = maxColumns;
        td.textContent = `... 还有 ${totalRows - maxDisplayRows} 行数据未显示（为提高性能）`;
        td.style.textAlign = 'center';
        td.style.fontStyle = 'italic';
        td.style.color = '#666';
        td.style.backgroundColor = '#f8f9fa';
        
        tr.appendChild(td);
        tableBody.appendChild(tr);
    }

    /**
     * 显示数据统计信息
     * @param {Array} data - 表格数据
     */
    function displayDataStats(data) {
        if (!dataStats) {
            return;
        }

        const totalRows = data.length - 1; // 减去表头行
        const totalColumns = data.length > 0 ? Math.max(...data.map(row => row.length)) : 0;
        const displayedRows = Math.min(totalRows, maxDisplayRows);
        
        dataStats.innerHTML = `
            <strong>数据统计：</strong>
            总行数: ${totalRows} | 
            总列数: ${totalColumns} | 
            显示行数: ${displayedRows}
            ${totalRows > maxDisplayRows ? ' | <span style="color: #dc3545;">数据已截断显示</span>' : ''}
        `;
    }

    /**
     * 隐藏数据展示区域
     */
    function hideDataDisplay() {
        if (dataDisplay) {
            dataDisplay.style.display = 'none';
        }
    }

    /**
     * 清空表格内容
     */
    function clearTable() {
        if (tableHead) {
            tableHead.innerHTML = '';
        }
        if (tableBody) {
            tableBody.innerHTML = '';
        }
        if (dataStats) {
            dataStats.innerHTML = '';
        }
    }

    /**
     * 重置模块状态
     */
    function reset() {
        // 隐藏相关区域
        if (sheetSelector) {
            sheetSelector.style.display = 'none';
        }
        hideDataDisplay();
        
        // 清空内容
        clearTable();
        
        if (sheetSelect) {
            sheetSelect.innerHTML = '<option value="">请选择工作表</option>';
        }
    }

    /**
     * 设置最大显示行数
     * @param {number} rows - 最大显示行数
     */
    function setMaxDisplayRows(rows) {
        if (typeof rows === 'number' && rows > 0) {
            maxDisplayRows = rows;
        }
    }

    /**
     * 获取当前选中的工作表名称
     * @returns {string} 工作表名称
     */
    function getCurrentSheet() {
        return sheetSelect ? sheetSelect.value : '';
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
        populateSheetSelector: populateSheetSelector,
        displaySheet: displaySheet,
        reset: reset,
        setMaxDisplayRows: setMaxDisplayRows,
        getCurrentSheet: getCurrentSheet,
        isReady: isReady
    };

})();

// 模块加载完成日志
console.log('TableRenderer模块已加载 - 表格渲染功能已就绪');
