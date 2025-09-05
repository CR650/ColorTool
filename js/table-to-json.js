/**
 * 表格转JSON工具 - 核心功能模块
 * 文件说明：实现表格文件解析、JSON转换和下载功能
 * 创建时间：2025-01-09
 * 功能：Excel/CSV文件处理、多种JSON格式转换、实时预览
 */

(function() {
    'use strict';

    // 全局变量
    let currentWorkbook = null;
    let currentData = null;
    let currentFileName = '';
    let convertedJson = null;

    // DOM元素引用
    const elements = {
        uploadArea: null,
        fileInput: null,
        selectFileBtn: null,
        uploadProgress: null,
        progressFill: null,
        progressText: null,
        fileInfo: null,
        fileName: null,
        fileSize: null,
        fileType: null,
        sheetCount: null,
        sheetSelection: null,
        sheetSelect: null,
        sheetInfo: null,
        conversionOptions: null,
        convertBtn: null,
        resetBtn: null,
        dataPreview: null,
        previewTable: null,
        previewTableHead: null,
        previewTableBody: null,
        jsonPreview: null,
        tableStats: null,
        jsonStats: null,
        downloadSection: null,
        downloadJsonBtn: null,
        copyJsonBtn: null,
        conversionStatus: null,
        dataRows: null,
        jsonSize: null,
        statusMessage: null
    };

    // 初始化函数
    function init() {
        console.log('表格转JSON工具初始化开始');
        
        // 检查依赖
        if (!window.XLSX) {
            showStatus('错误：XLSX库未加载', 'error');
            return;
        }

        // 获取DOM元素
        initializeElements();
        
        // 设置事件监听器
        setupEventListeners();
        
        console.log('表格转JSON工具初始化完成');
        showStatus('工具已就绪，请选择表格文件', 'info');
    }

    // 初始化DOM元素引用
    function initializeElements() {
        const elementIds = [
            'uploadArea', 'fileInput', 'selectFileBtn', 'uploadProgress', 
            'progressFill', 'progressText', 'fileInfo', 'fileName', 'fileSize', 
            'fileType', 'sheetCount', 'sheetSelection', 'sheetSelect', 'sheetInfo',
            'conversionOptions', 'convertBtn', 'resetBtn', 'dataPreview',
            'previewTable', 'previewTableHead', 'previewTableBody', 'jsonPreview',
            'tableStats', 'jsonStats', 'downloadSection', 'downloadJsonBtn',
            'copyJsonBtn', 'conversionStatus', 'dataRows', 'jsonSize', 'statusMessage'
        ];

        elementIds.forEach(id => {
            elements[id] = document.getElementById(id);
            if (!elements[id]) {
                console.warn(`元素未找到: ${id}`);
            }
        });
    }

    // 设置事件监听器
    function setupEventListeners() {
        // 文件选择
        if (elements.selectFileBtn) {
            elements.selectFileBtn.addEventListener('click', () => {
                elements.fileInput.click();
            });
        }

        if (elements.fileInput) {
            elements.fileInput.addEventListener('change', handleFileSelect);
        }

        // 拖拽上传
        if (elements.uploadArea) {
            elements.uploadArea.addEventListener('dragover', handleDragOver);
            elements.uploadArea.addEventListener('dragleave', handleDragLeave);
            elements.uploadArea.addEventListener('drop', handleFileDrop);
            elements.uploadArea.addEventListener('click', () => {
                elements.fileInput.click();
            });
        }

        // 工作表选择
        if (elements.sheetSelect) {
            elements.sheetSelect.addEventListener('change', handleSheetChange);
        }

        // 转换选项变化
        const optionInputs = document.querySelectorAll('#conversionOptions input');
        optionInputs.forEach(input => {
            input.addEventListener('change', updatePreview);
        });

        // 按钮事件
        if (elements.convertBtn) {
            elements.convertBtn.addEventListener('click', convertToJson);
        }

        if (elements.resetBtn) {
            elements.resetBtn.addEventListener('click', resetTool);
        }

        if (elements.downloadJsonBtn) {
            elements.downloadJsonBtn.addEventListener('click', downloadJson);
        }

        if (elements.copyJsonBtn) {
            elements.copyJsonBtn.addEventListener('click', copyJsonToClipboard);
        }

        // 预览标签切换
        const tabBtns = document.querySelectorAll('.tab-btn');
        tabBtns.forEach(btn => {
            btn.addEventListener('click', (e) => {
                switchTab(e.target.dataset.tab);
            });
        });
    }

    // 处理文件选择
    function handleFileSelect(event) {
        const file = event.target.files[0];
        if (file) {
            processFile(file);
        }
    }

    // 处理拖拽悬停
    function handleDragOver(event) {
        event.preventDefault();
        elements.uploadArea.classList.add('dragover');
    }

    // 处理拖拽离开
    function handleDragLeave(event) {
        event.preventDefault();
        elements.uploadArea.classList.remove('dragover');
    }

    // 处理文件拖拽
    function handleFileDrop(event) {
        event.preventDefault();
        elements.uploadArea.classList.remove('dragover');
        
        const files = event.dataTransfer.files;
        if (files.length > 0) {
            processFile(files[0]);
        }
    }

    // 处理文件
    function processFile(file) {
        console.log('开始处理文件:', file.name);
        
        // 验证文件类型
        if (!validateFileType(file)) {
            return;
        }

        currentFileName = file.name;
        
        // 显示进度
        showProgress(true);
        updateProgress(0, '正在读取文件...');

        // 读取文件
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                updateProgress(30, '正在解析文件...');
                
                if (file.name.toLowerCase().endsWith('.csv')) {
                    parseCSVFile(e.target.result);
                } else {
                    parseExcelFile(e.target.result);
                }
                
                updateProgress(100, '文件处理完成');
                setTimeout(() => showProgress(false), 1000);
                
            } catch (error) {
                console.error('文件处理错误:', error);
                showStatus(`文件处理失败: ${error.message}`, 'error');
                showProgress(false);
            }
        };

        reader.onerror = function() {
            showStatus('文件读取失败', 'error');
            showProgress(false);
        };

        // 根据文件类型选择读取方式
        if (file.name.toLowerCase().endsWith('.csv')) {
            reader.readAsText(file, 'UTF-8');
        } else {
            reader.readAsArrayBuffer(file);
        }
    }

    // 验证文件类型
    function validateFileType(file) {
        const allowedTypes = ['.xlsx', '.xls', '.csv'];
        const fileName = file.name.toLowerCase();
        const isValid = allowedTypes.some(type => fileName.endsWith(type));
        
        if (!isValid) {
            showStatus('不支持的文件格式，请选择 Excel (.xlsx, .xls) 或 CSV (.csv) 文件', 'error');
            return false;
        }
        
        // 检查文件大小 (限制为50MB)
        const maxSize = 50 * 1024 * 1024;
        if (file.size > maxSize) {
            showStatus('文件过大，请选择小于50MB的文件', 'error');
            return false;
        }
        
        return true;
    }

    // 解析CSV文件
    function parseCSVFile(csvText) {
        console.log('解析CSV文件');
        
        const lines = csvText.split('\n').filter(line => line.trim());
        if (lines.length === 0) {
            throw new Error('CSV文件为空');
        }

        const data = lines.map(line => parseCSVLine(line));
        
        // 创建模拟的工作簿结构
        currentWorkbook = {
            SheetNames: ['Sheet1'],
            Sheets: {
                'Sheet1': data
            }
        };
        
        currentData = { 'Sheet1': data };
        
        // 更新UI
        updateFileInfo(currentFileName, csvText.length, 'CSV', 1);
        showSheetSelection(['Sheet1']);
        
        showStatus('CSV文件解析成功', 'success');
    }

    // 解析CSV行
    function parseCSVLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        
        result.push(current.trim());
        return result;
    }

    // 解析Excel文件
    function parseExcelFile(arrayBuffer) {
        console.log('解析Excel文件');
        
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

        // 更新UI
        updateFileInfo(currentFileName, arrayBuffer.byteLength, 'Excel', workbook.SheetNames.length);
        showSheetSelection(workbook.SheetNames);
        
        showStatus(`Excel文件解析成功，包含 ${workbook.SheetNames.length} 个工作表`, 'success');
    }

    // 更新文件信息显示
    function updateFileInfo(fileName, fileSize, fileType, sheetCount) {
        if (elements.fileName) elements.fileName.textContent = fileName;
        if (elements.fileSize) elements.fileSize.textContent = formatFileSize(fileSize);
        if (elements.fileType) elements.fileType.textContent = fileType;
        if (elements.sheetCount) elements.sheetCount.textContent = sheetCount;

        if (elements.fileInfo) {
            elements.fileInfo.style.display = 'block';
        }
    }

    // 格式化文件大小
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    // 显示工作表选择
    function showSheetSelection(sheetNames) {
        if (!elements.sheetSelect) return;

        // 清空现有选项
        elements.sheetSelect.innerHTML = '<option value="">请选择工作表</option>';

        // 添加工作表选项
        sheetNames.forEach(sheetName => {
            const option = document.createElement('option');
            option.value = sheetName;
            option.textContent = sheetName;
            elements.sheetSelect.appendChild(option);
        });

        // 如果只有一个工作表，自动选择
        if (sheetNames.length === 1) {
            elements.sheetSelect.value = sheetNames[0];
            handleSheetChange();
        }

        if (elements.sheetSelection) {
            elements.sheetSelection.style.display = 'block';
        }
    }

    // 处理工作表选择变化
    function handleSheetChange() {
        const selectedSheet = elements.sheetSelect.value;
        if (!selectedSheet || !currentData[selectedSheet]) {
            return;
        }

        const sheetData = currentData[selectedSheet];
        const rowCount = sheetData.length;
        const colCount = rowCount > 0 ? Math.max(...sheetData.map(row => row.length)) : 0;

        // 更新工作表信息
        if (elements.sheetInfo) {
            elements.sheetInfo.textContent = `${rowCount} 行 × ${colCount} 列`;
        }

        // 显示转换选项和预览
        showConversionOptions();
        updatePreview();

        showStatus(`已选择工作表: ${selectedSheet}`, 'info');
    }

    // 显示转换选项
    function showConversionOptions() {
        if (elements.conversionOptions) {
            elements.conversionOptions.style.display = 'block';
        }

        // 启用转换按钮
        if (elements.convertBtn) {
            elements.convertBtn.disabled = false;
        }
    }

    // 更新预览
    function updatePreview() {
        const selectedSheet = elements.sheetSelect.value;
        if (!selectedSheet || !currentData[selectedSheet]) {
            return;
        }

        const sheetData = currentData[selectedSheet];
        const options = getConversionOptions();

        // 处理数据
        const processedData = processDataWithOptions(sheetData, options);

        // 更新表格预览
        updateTablePreview(processedData);

        // 更新JSON预览
        const jsonData = convertDataToJson(processedData, options);
        updateJsonPreview(jsonData);

        // 显示预览区域
        if (elements.dataPreview) {
            elements.dataPreview.style.display = 'block';
        }
    }

    // 获取转换选项
    function getConversionOptions() {
        return {
            includeHeader: document.getElementById('includeHeader')?.checked || false,
            skipEmptyRows: document.getElementById('skipEmptyRows')?.checked || false,
            trimWhitespace: document.getElementById('trimWhitespace')?.checked || false,
            jsonFormat: document.querySelector('input[name="jsonFormat"]:checked')?.value || 'array',
            emptyValue: document.querySelector('input[name="emptyValue"]:checked')?.value || 'null'
        };
    }

    // 根据选项处理数据
    function processDataWithOptions(data, options) {
        let processedData = [...data];

        // 跳过空行
        if (options.skipEmptyRows) {
            processedData = processedData.filter(row =>
                row.some(cell => cell !== null && cell !== undefined && cell.toString().trim() !== '')
            );
        }

        // 去除首尾空格
        if (options.trimWhitespace) {
            processedData = processedData.map(row =>
                row.map(cell =>
                    typeof cell === 'string' ? cell.trim() : cell
                )
            );
        }

        return processedData;
    }

    // 更新表格预览
    function updateTablePreview(data) {
        if (!elements.previewTableHead || !elements.previewTableBody) return;

        // 清空现有内容
        elements.previewTableHead.innerHTML = '';
        elements.previewTableBody.innerHTML = '';

        if (data.length === 0) return;

        // 限制预览行数
        const maxPreviewRows = 100;
        const previewData = data.slice(0, maxPreviewRows + 1); // +1 for header

        // 创建表头
        if (previewData.length > 0) {
            const headerRow = document.createElement('tr');
            const headers = previewData[0];

            headers.forEach((header, index) => {
                const th = document.createElement('th');
                th.textContent = header || `列${index + 1}`;
                headerRow.appendChild(th);
            });

            elements.previewTableHead.appendChild(headerRow);
        }

        // 创建表格内容
        const bodyData = previewData.slice(1);
        bodyData.forEach((row, rowIndex) => {
            if (rowIndex >= maxPreviewRows) return;

            const tr = document.createElement('tr');
            row.forEach(cell => {
                const td = document.createElement('td');
                td.textContent = cell || '';
                tr.appendChild(td);
            });
            elements.previewTableBody.appendChild(tr);
        });

        // 更新统计信息
        if (elements.tableStats) {
            const totalRows = data.length - 1; // 减去表头
            const totalCols = data.length > 0 ? data[0].length : 0;
            const showingRows = Math.min(totalRows, maxPreviewRows);

            elements.tableStats.textContent =
                `显示 ${showingRows} / ${totalRows} 行，共 ${totalCols} 列`;
        }
    }

    // 将数据转换为JSON
    function convertDataToJson(data, options) {
        if (data.length === 0) return null;

        let headers = [];
        let rows = [];

        if (options.includeHeader && data.length > 0) {
            headers = data[0];
            rows = data.slice(1);
        } else {
            // 生成默认列名
            const maxCols = Math.max(...data.map(row => row.length));
            headers = Array.from({length: maxCols}, (_, i) => `列${i + 1}`);
            rows = data;
        }

        // 根据格式转换
        switch (options.jsonFormat) {
            case 'array':
                return convertToObjectArray(headers, rows, options);
            case 'keyValue':
                return convertToKeyValue(headers, rows, options);
            case 'nested':
                return convertToNested(headers, rows, options);
            default:
                return convertToObjectArray(headers, rows, options);
        }
    }

    // 转换为对象数组格式
    function convertToObjectArray(headers, rows, options) {
        return rows.map((row, index) => {
            const obj = {};
            headers.forEach((header, colIndex) => {
                const value = row[colIndex];
                const processedValue = processValue(value, options);

                if (options.emptyValue === 'skip' && (processedValue === null || processedValue === '')) {
                    return; // 跳过空值
                }

                obj[header || `列${colIndex + 1}`] = processedValue;
            });
            return obj;
        });
    }

    // 转换为键值对格式
    function convertToKeyValue(headers, rows, options) {
        const result = {};
        rows.forEach((row, rowIndex) => {
            const rowKey = `行${rowIndex + 1}`;
            const obj = {};
            headers.forEach((header, colIndex) => {
                const value = row[colIndex];
                const processedValue = processValue(value, options);

                if (options.emptyValue === 'skip' && (processedValue === null || processedValue === '')) {
                    return;
                }

                obj[header || `列${colIndex + 1}`] = processedValue;
            });
            result[rowKey] = obj;
        });
        return result;
    }

    // 转换为嵌套对象格式
    function convertToNested(headers, rows, options) {
        const data = convertToObjectArray(headers, rows, options);
        return {
            metadata: {
                fileName: currentFileName,
                sheetName: elements.sheetSelect.value,
                totalRows: rows.length,
                totalColumns: headers.length,
                exportTime: new Date().toISOString(),
                options: options
            },
            headers: headers,
            data: data
        };
    }

    // 处理值
    function processValue(value, options) {
        if (value === null || value === undefined || value === '') {
            switch (options.emptyValue) {
                case 'null':
                    return null;
                case 'empty':
                    return '';
                case 'skip':
                    return null; // 将在上层处理跳过
                default:
                    return null;
            }
        }

        // 尝试转换数字
        if (typeof value === 'string' && !isNaN(value) && !isNaN(parseFloat(value))) {
            const num = parseFloat(value);
            return Number.isInteger(num) ? parseInt(value) : num;
        }

        return value;
    }

    // 更新JSON预览
    function updateJsonPreview(jsonData) {
        if (!elements.jsonPreview) return;

        if (jsonData) {
            const jsonString = JSON.stringify(jsonData, null, 2);
            elements.jsonPreview.textContent = jsonString;

            // 更新统计信息
            if (elements.jsonStats) {
                const size = new Blob([jsonString]).size;
                const itemCount = Array.isArray(jsonData) ? jsonData.length :
                                 (jsonData.data ? jsonData.data.length : Object.keys(jsonData).length);

                elements.jsonStats.textContent =
                    `JSON大小: ${formatFileSize(size)}，包含 ${itemCount} 个数据项`;
            }
        } else {
            elements.jsonPreview.textContent = '暂无数据';
            if (elements.jsonStats) {
                elements.jsonStats.textContent = '';
            }
        }
    }

    // 转换为JSON
    function convertToJson() {
        const selectedSheet = elements.sheetSelect.value;
        if (!selectedSheet || !currentData[selectedSheet]) {
            showStatus('请先选择工作表', 'warning');
            return;
        }

        try {
            const sheetData = currentData[selectedSheet];
            const options = getConversionOptions();

            // 处理数据
            const processedData = processDataWithOptions(sheetData, options);

            // 转换为JSON
            convertedJson = convertDataToJson(processedData, options);

            if (!convertedJson) {
                showStatus('转换失败：数据为空', 'error');
                return;
            }

            // 更新下载区域
            updateDownloadSection();

            // 更新预览
            updateJsonPreview(convertedJson);

            showStatus('JSON转换成功', 'success');

        } catch (error) {
            console.error('转换错误:', error);
            showStatus(`转换失败: ${error.message}`, 'error');
        }
    }

    // 更新下载区域
    function updateDownloadSection() {
        if (!convertedJson) return;

        const jsonString = JSON.stringify(convertedJson, null, 2);
        const size = new Blob([jsonString]).size;
        const itemCount = Array.isArray(convertedJson) ? convertedJson.length :
                         (convertedJson.data ? convertedJson.data.length : Object.keys(convertedJson).length);

        // 更新状态信息
        if (elements.conversionStatus) elements.conversionStatus.textContent = '转换成功';
        if (elements.dataRows) elements.dataRows.textContent = itemCount;
        if (elements.jsonSize) elements.jsonSize.textContent = formatFileSize(size);

        // 启用下载按钮
        if (elements.downloadJsonBtn) elements.downloadJsonBtn.disabled = false;
        if (elements.copyJsonBtn) elements.copyJsonBtn.disabled = false;

        // 显示下载区域
        if (elements.downloadSection) {
            elements.downloadSection.style.display = 'block';
        }
    }

    // 下载JSON文件
    function downloadJson() {
        if (!convertedJson) {
            showStatus('没有可下载的数据', 'warning');
            return;
        }

        try {
            const jsonString = JSON.stringify(convertedJson, null, 2);
            const blob = new Blob([jsonString], { type: 'application/json;charset=utf-8' });

            // 生成文件名
            const baseName = currentFileName.replace(/\.[^/.]+$/, '');
            const sheetName = elements.sheetSelect.value;
            const fileName = `${baseName}_${sheetName}.json`;

            // 创建下载链接
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            showStatus('JSON文件下载成功', 'success');

        } catch (error) {
            console.error('下载错误:', error);
            showStatus(`下载失败: ${error.message}`, 'error');
        }
    }

    // 复制JSON到剪贴板
    function copyJsonToClipboard() {
        if (!convertedJson) {
            showStatus('没有可复制的数据', 'warning');
            return;
        }

        try {
            const jsonString = JSON.stringify(convertedJson, null, 2);

            if (navigator.clipboard) {
                navigator.clipboard.writeText(jsonString).then(() => {
                    showStatus('JSON已复制到剪贴板', 'success');
                }).catch(error => {
                    console.error('复制失败:', error);
                    fallbackCopyToClipboard(jsonString);
                });
            } else {
                fallbackCopyToClipboard(jsonString);
            }

        } catch (error) {
            console.error('复制错误:', error);
            showStatus(`复制失败: ${error.message}`, 'error');
        }
    }

    // 备用复制方法
    function fallbackCopyToClipboard(text) {
        const textArea = document.createElement('textarea');
        textArea.value = text;
        textArea.style.position = 'fixed';
        textArea.style.left = '-999999px';
        textArea.style.top = '-999999px';
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();

        try {
            document.execCommand('copy');
            showStatus('JSON已复制到剪贴板', 'success');
        } catch (error) {
            showStatus('复制失败，请手动复制', 'error');
        }

        document.body.removeChild(textArea);
    }

    // 切换预览标签
    function switchTab(tabName) {
        // 更新标签按钮状态
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        // 更新标签内容
        document.querySelectorAll('.tab-pane').forEach(pane => {
            pane.classList.remove('active');
        });

        if (tabName === 'table') {
            document.getElementById('tableTab').classList.add('active');
        } else if (tabName === 'json') {
            document.getElementById('jsonTab').classList.add('active');
        }
    }

    // 重置工具
    function resetTool() {
        // 重置变量
        currentWorkbook = null;
        currentData = null;
        currentFileName = '';
        convertedJson = null;

        // 重置文件输入
        if (elements.fileInput) {
            elements.fileInput.value = '';
        }

        // 隐藏所有区域
        const sectionsToHide = [
            'fileInfo', 'sheetSelection', 'conversionOptions',
            'dataPreview', 'downloadSection'
        ];

        sectionsToHide.forEach(sectionId => {
            const element = elements[sectionId];
            if (element) {
                element.style.display = 'none';
            }
        });

        // 重置按钮状态
        if (elements.convertBtn) elements.convertBtn.disabled = true;
        if (elements.downloadJsonBtn) elements.downloadJsonBtn.disabled = true;
        if (elements.copyJsonBtn) elements.copyJsonBtn.disabled = true;

        // 清空内容
        if (elements.previewTableHead) elements.previewTableHead.innerHTML = '';
        if (elements.previewTableBody) elements.previewTableBody.innerHTML = '';
        if (elements.jsonPreview) elements.jsonPreview.textContent = '';

        // 重置拖拽区域
        if (elements.uploadArea) {
            elements.uploadArea.classList.remove('dragover');
        }

        showStatus('工具已重置', 'info');
    }

    // 显示进度
    function showProgress(show) {
        if (elements.uploadProgress) {
            elements.uploadProgress.style.display = show ? 'block' : 'none';
        }
    }

    // 更新进度
    function updateProgress(percent, text) {
        if (elements.progressFill) {
            elements.progressFill.style.width = percent + '%';
        }
        if (elements.progressText) {
            elements.progressText.textContent = text;
        }
    }

    // 显示状态消息
    function showStatus(message, type = 'info') {
        if (!elements.statusMessage) return;

        // 清除现有类
        elements.statusMessage.className = 'status-message';

        // 添加类型类
        elements.statusMessage.classList.add(type);
        elements.statusMessage.textContent = message;

        // 显示消息
        elements.statusMessage.classList.add('show');

        // 自动隐藏
        setTimeout(() => {
            elements.statusMessage.classList.remove('show');
        }, type === 'error' ? 5000 : 3000);

        console.log(`[${type.toUpperCase()}] ${message}`);
    }

    // 工具函数：防抖
    function debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    }

    // 工具函数：节流
    function throttle(func, limit) {
        let inThrottle;
        return function() {
            const args = arguments;
            const context = this;
            if (!inThrottle) {
                func.apply(context, args);
                inThrottle = true;
                setTimeout(() => inThrottle = false, limit);
            }
        };
    }

    // 错误处理
    window.addEventListener('error', function(event) {
        console.error('全局错误:', event.error);
        showStatus('发生未知错误，请刷新页面重试', 'error');
    });

    // 页面卸载前的清理
    window.addEventListener('beforeunload', function() {
        // 清理对象URL
        if (convertedJson) {
            // 清理可能的URL对象
        }
    });

    // 页面加载完成后初始化
    document.addEventListener('DOMContentLoaded', function() {
        console.log('DOM加载完成，开始初始化表格转JSON工具');

        // 延迟初始化以确保所有资源加载完成
        setTimeout(init, 100);
    });

    // 暴露公共接口（用于调试）
    window.TableToJsonTool = {
        init: init,
        reset: resetTool,
        getCurrentData: () => currentData,
        getCurrentJson: () => convertedJson,
        showStatus: showStatus
    };

    console.log('表格转JSON工具模块已加载');

})();
