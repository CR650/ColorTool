/**
 * 文件处理模块
 * 文件说明：负责文件选择、拖拽上传、文件验证等功能
 * 创建时间：2025-01-09
 * 依赖：App.Utils, App.DataParser
 */

// 确保全局命名空间存在
window.App = window.App || {};

/**
 * 文件处理模块
 * 处理文件上传、拖拽、验证等相关功能
 */
window.App.FileHandler = (function() {
    'use strict';

    // DOM元素引用
    let sourceUploadArea = null;
    let sourceFileInput = null;

    // 模块状态
    let isInitialized = false;

    /**
     * 初始化文件处理模块
     * 获取DOM元素引用并设置事件监听器
     */
    function init() {
        if (isInitialized) {
            console.warn('FileHandler模块已经初始化');
            return;
        }

        // 获取DOM元素
        sourceUploadArea = document.getElementById('sourceUploadArea');
        sourceFileInput = document.getElementById('sourceFileInput');

        // 验证必要元素是否存在
        if (!sourceUploadArea || !sourceFileInput) {
            console.error('FileHandler初始化失败：缺少源文件上传元素');
            return;
        }

        // 设置事件监听器
        setupEventListeners();
        
        isInitialized = true;
        console.log('FileHandler模块初始化完成');
    }

    /**
     * 设置所有事件监听器
     * 包括文件选择、拖拽事件等
     */
    function setupEventListeners() {
        // 源数据文件相关事件
        sourceFileInput.addEventListener('change', handleSourceFileSelect);
        sourceUploadArea.addEventListener('dragover', (e) => handleDragOver(e, 'source'));
        sourceUploadArea.addEventListener('dragleave', (e) => handleDragLeave(e, 'source'));
        sourceUploadArea.addEventListener('drop', (e) => handleFileDrop(e, 'source'));
        sourceUploadArea.addEventListener('click', () => {
            // 显示上次选择的文件提示
            const lastFileInfo = App.Utils.getLastPath('SOURCE_DATA');
            if (lastFileInfo && lastFileInfo.fileName) {
                console.log(`上次选择的源数据文件: ${lastFileInfo.fileName}`);
                App.Utils.showStatus(`提示：上次选择的文件是 ${lastFileInfo.fileName}`, 'info', 2000);
            }
            sourceFileInput.click();
        });

        // 阻止默认拖拽行为
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            document.addEventListener(eventName, preventDefaults, false);
        });
    }

    /**
     * 阻止默认事件行为
     * 防止浏览器默认的拖拽处理
     * @param {Event} e - 事件对象
     */
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    /**
     * 处理拖拽悬停事件
     * @param {DragEvent} e - 拖拽事件对象
     * @param {string} type - 上传区域类型
     */
    function handleDragOver(e, type) {
        e.preventDefault();
        if (type === 'source') {
            sourceUploadArea.classList.add('dragover');
        }
    }

    /**
     * 处理拖拽离开事件
     * @param {DragEvent} e - 拖拽事件对象
     * @param {string} type - 上传区域类型
     */
    function handleDragLeave(e, type) {
        e.preventDefault();
        if (type === 'source') {
            sourceUploadArea.classList.remove('dragover');
        }
    }

    /**
     * 处理文件拖拽放置事件
     * @param {DragEvent} e - 拖拽事件对象
     * @param {string} type - 上传区域类型
     */
    function handleFileDrop(e, type) {
        e.preventDefault();
        if (type === 'source') {
            sourceUploadArea.classList.remove('dragover');
        }

        const files = e.dataTransfer.files;
        if (files.length > 0) {
            if (type === 'source') {
                processSourceFile(files[0]);
            }
        }
    }

    /**
     * 处理源数据文件选择事件
     * @param {Event} e - 文件选择事件对象
     */
    function handleSourceFileSelect(e) {
        const files = e.target.files;
        if (files.length > 0) {
            processSourceFile(files[0]);
        }
    }



    /**
     * 处理源数据文件
     * @param {File} file - 源数据文件对象
     */
    function processSourceFile(file) {
        // 验证文件类型（只支持Excel文件）
        const allowedTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];

        if (!allowedTypes.includes(file.type) && !file.name.toLowerCase().match(/\.(xlsx|xls)$/)) {
            App.Utils.showStatus('源数据文件必须是Excel格式 (.xlsx, .xls)', 'error');

            // 隐藏文件选择结果显示
            hideSourceFileSelectionResult();

            // 显示错误状态
            if (window.App && window.App.ThemeManager && window.App.ThemeManager.updateFileSelectionStatus) {
                window.App.ThemeManager.updateFileSelectionStatus('sourceFileStatus', 'error', {
                    fileName: file.name,
                    fileSize: file.size,
                    errorMessage: '文件格式不支持，请选择Excel文件(.xlsx或.xls)'
                });
            }
            return;
        }

        // 验证文件大小
        if (!App.Utils.validateFileSize(file, 10)) {
            App.Utils.showStatus('文件过大！请选择小于10MB的文件', 'error');

            // 隐藏文件选择结果显示
            hideSourceFileSelectionResult();

            // 显示错误状态
            if (window.App && window.App.ThemeManager && window.App.ThemeManager.updateFileSelectionStatus) {
                window.App.ThemeManager.updateFileSelectionStatus('sourceFileStatus', 'error', {
                    fileName: file.name,
                    fileSize: file.size,
                    errorMessage: '文件过大，请选择小于10MB的文件'
                });
            }
            return;
        }

        // 显示加载状态
        if (window.App && window.App.ThemeManager && window.App.ThemeManager.updateFileSelectionStatus) {
            window.App.ThemeManager.updateFileSelectionStatus('sourceFileStatus', 'loading', {
                fileName: file.name,
                fileSize: file.size
            });
        }

        // 保存文件记忆
        if (file.name) {
            App.Utils.saveLastPath('SOURCE_DATA', file.name);
        }

        App.Utils.showStatus('正在读取源数据文件...', 'info');

        // 读取源数据文件
        readSourceFile(file);
    }



    /**
     * 读取源数据文件
     * @param {File} file - 源数据文件对象
     */
    function readSourceFile(file) {
        const reader = new FileReader();

        reader.onload = function(e) {
            try {
                // 解析Excel文件
                const workbook = XLSX.read(e.target.result, { type: 'array' });

                // 查找"完整配色表"工作表
                const targetSheetName = '完整配色表';
                let worksheet = null;

                console.log('可用工作表:', workbook.SheetNames);

                if (workbook.SheetNames.includes(targetSheetName)) {
                    worksheet = workbook.Sheets[targetSheetName];
                    console.log(`找到目标工作表: ${targetSheetName}`);
                } else {
                    // 如果没有找到"完整配色表"工作表，尝试查找包含"配色"或"颜色"的工作表
                    const colorSheetNames = workbook.SheetNames.filter(name =>
                        name.includes('配色') || name.includes('颜色') || name.includes('color') || name.includes('Color')
                    );

                    if (colorSheetNames.length > 0) {
                        worksheet = workbook.Sheets[colorSheetNames[0]];
                        console.log(`未找到"完整配色表"工作表，使用: ${colorSheetNames[0]}`);
                        App.Utils.showStatus(`警告：未找到"完整配色表"工作表，使用"${colorSheetNames[0]}"工作表`, 'warning');
                    } else {
                        // 最后使用第一个工作表
                        worksheet = workbook.Sheets[workbook.SheetNames[0]];
                        console.log(`未找到配色相关工作表，使用第一个工作表: ${workbook.SheetNames[0]}`);
                        App.Utils.showStatus(`警告：未找到"完整配色表"工作表，使用第一个工作表"${workbook.SheetNames[0]}"`, 'warning');
                    }
                }

                const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                    header: 1,
                    defval: '',
                    raw: false
                });

                // 转换为对象数组格式
                const headers = jsonData[0] || [];
                console.log('Excel文件原始表头:', headers);

                // 检测字段兼容性（改为警告模式，不阻止文件加载）
                const colorCodeFields = ['颜色代码', 'colorCode', 'code', '代码', 'Color Code', 'ColorCode'];
                const colorValueFields = ['16进制值', '颜色值', 'hex', 'HEX', 'hexValue', '16进制', 'color', 'Color', 'Hex Value'];
                const rgbFields = ['R值', 'G值', 'B值'];
                const rgbFieldsAlt = ['R', 'G', 'B'];
                const rgbFieldsEng = ['red', 'green', 'blue'];

                const hasColorCode = colorCodeFields.some(field => headers.includes(field));
                const hasColorValue = colorValueFields.some(field => headers.includes(field));
                const hasRGB = rgbFields.every(field => headers.includes(field));
                const hasRGBAlt = rgbFieldsAlt.every(field => headers.includes(field));
                const hasRGBEng = rgbFieldsEng.every(field => headers.includes(field));
                const hasAnyRGB = hasRGB || hasRGBAlt || hasRGBEng;

                console.log('字段兼容性检测结果:', {
                    hasColorCode,
                    hasColorValue,
                    hasRGB,
                    hasRGBAlt,
                    hasRGBEng,
                    hasAnyRGB,
                    availableFields: headers.filter(h => h && h.trim()),
                    detectedColorCodeField: colorCodeFields.find(field => headers.includes(field)),
                    detectedColorValueField: colorValueFields.find(field => headers.includes(field))
                });

                // 改为警告模式，不阻止文件加载
                if (!hasColorCode) {
                    console.warn('⚠️ 未检测到标准的颜色代码字段，将尝试使用第一个可用字段');
                    App.Utils.showStatus('警告：未检测到标准颜色代码字段，但将继续尝试处理', 'warning');
                }

                if (!hasColorValue && !hasAnyRGB) {
                    console.warn('⚠️ 未检测到标准的颜色值字段，颜色提取可能受影响');
                    App.Utils.showStatus('警告：未检测到标准颜色值字段，部分颜色可能使用默认值', 'warning');
                }

                const data = jsonData.slice(1).map((row, index) => {
                    const obj = {};
                    headers.forEach((header, headerIndex) => {
                        obj[header] = row[headerIndex] || '';
                    });

                    // 调试：输出前3行数据
                    if (index < 3) {
                        console.log(`第${index + 1}行数据:`, obj);
                    }

                    return obj;
                });

                const sourceData = {
                    headers: headers,
                    data: data,
                    fileName: file.name,
                    fileSize: file.size, // 添加文件大小信息
                    workbook: workbook, // 添加workbook对象，用于映射模式检测
                    validation: {
                        hasColorCode,
                        hasColorValue,
                        hasRGB,
                        totalRows: data.length
                    }
                };

                // 传递给主题管理器
                if (App.ThemeManager) {
                    App.ThemeManager.setSourceData(sourceData);
                    App.Utils.showStatus('源数据文件读取成功', 'success');

                    // 显示文件选择结果
                    updateSourceFileSelectionResult(file, sourceData);

                    // 显示成功状态（保持兼容性）
                    if (window.App && window.App.ThemeManager && window.App.ThemeManager.updateFileSelectionStatus) {
                        window.App.ThemeManager.updateFileSelectionStatus('sourceFileStatus', 'success', {
                            fileName: file.name,
                            fileSize: file.size
                        });
                    }
                } else {
                    App.Utils.showStatus('主题管理模块未加载', 'error');

                    // 隐藏文件选择结果显示
                    hideSourceFileSelectionResult();

                    // 显示错误状态
                    if (window.App && window.App.ThemeManager && window.App.ThemeManager.updateFileSelectionStatus) {
                        window.App.ThemeManager.updateFileSelectionStatus('sourceFileStatus', 'error', {
                            fileName: file.name,
                            fileSize: file.size,
                            errorMessage: '主题管理模块未加载'
                        });
                    }
                }
            } catch (error) {
                console.error('源数据文件解析错误:', error);
                App.Utils.showStatus('源数据文件解析失败：' + error.message, 'error');

                // 隐藏文件选择结果显示
                hideSourceFileSelectionResult();

                // 显示错误状态
                if (window.App && window.App.ThemeManager && window.App.ThemeManager.updateFileSelectionStatus) {
                    window.App.ThemeManager.updateFileSelectionStatus('sourceFileStatus', 'error', {
                        fileName: file.name,
                        fileSize: file.size,
                        errorMessage: '文件解析失败：' + error.message
                    });
                }
            }
        };

        reader.onerror = function() {
            App.Utils.showStatus('源数据文件读取失败', 'error');

            // 隐藏文件选择结果显示
            hideSourceFileSelectionResult();

            // 显示错误状态
            if (window.App && window.App.ThemeManager && window.App.ThemeManager.updateFileSelectionStatus) {
                window.App.ThemeManager.updateFileSelectionStatus('sourceFileStatus', 'error', {
                    fileName: file.name,
                    fileSize: file.size,
                    errorMessage: '文件读取失败'
                });
            }
        };

        reader.readAsArrayBuffer(file);
    }



    /**
     * 更新源数据文件选择结果显示
     * @param {File} file - 选择的文件对象
     * @param {Object} sourceData - 解析后的源数据
     */
    function updateSourceFileSelectionResult(file, sourceData) {
        const resultContainer = document.getElementById('sourceFileSelectionResult');
        const fileNameElement = document.getElementById('sourceFileName');
        const fileSizeElement = document.getElementById('sourceFileSize');
        const fileTimeElement = document.getElementById('sourceFileTime');
        const fileRowsElement = document.getElementById('sourceFileRows');

        if (!resultContainer) {
            console.warn('源数据文件选择结果容器未找到');
            return;
        }

        // 格式化文件大小
        function formatFileSize(bytes) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const sizes = ['Bytes', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }

        // 获取当前时间
        function getCurrentTimeString() {
            const now = new Date();
            return now.toLocaleString('zh-CN', {
                year: 'numeric',
                month: '2-digit',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            });
        }

        // 更新显示内容
        if (fileNameElement) {
            fileNameElement.textContent = file.name;
        }
        if (fileSizeElement) {
            fileSizeElement.textContent = formatFileSize(file.size);
        }
        if (fileTimeElement) {
            fileTimeElement.textContent = getCurrentTimeString();
        }
        if (fileRowsElement) {
            const rowCount = sourceData.data ? sourceData.data.length : 0;
            fileRowsElement.textContent = `${rowCount} 行`;
        }

        // 显示结果容器
        resultContainer.style.display = 'block';
    }

    /**
     * 隐藏源数据文件选择结果显示
     */
    function hideSourceFileSelectionResult() {
        const resultContainer = document.getElementById('sourceFileSelectionResult');
        if (resultContainer) {
            resultContainer.style.display = 'none';
        }
    }

    /**
     * 重置文件处理状态
     */
    function reset() {
        if (sourceFileInput) {
            sourceFileInput.value = '';
        }

        if (sourceUploadArea) {
            sourceUploadArea.classList.remove('dragover');
        }

        // 隐藏文件选择结果显示
        hideSourceFileSelectionResult();
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
        reset: reset,
        isReady: isReady
    };

})();

// 模块加载完成日志
console.log('FileHandler模块已加载 - 文件处理功能已就绪');
