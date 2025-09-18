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
    let rscLanguageData = null;      // RSC_Language.xls文件数据
    let mappingData = null;          // 对比映射数据
    let processedResult = null;      // 处理结果
    let rscAllSheetsData = null;     // RSC_Theme文件的所有Sheet数据
    let ugcAllSheetsData = null;     // UGCTheme文件的所有Sheet数据
    let multiLangConfig = null;      // 多语言配置数据

    // 文件夹选择相关状态
    let folderManager = null;        // Unity项目文件夹管理器实例
    let folderSelectionActive = false; // 是否使用文件夹选择模式

    // DOM元素引用
    let themeNameInput = null;
    let themeSelector = null;
    let processThemeBtn = null;
    let resetBtn = null;
    let enableDirectSaveBtn = null;
    let enableUGCDirectSaveBtn = null;
    let fileStatus = null;
    let operationModeIndicator = null;
    let modeBadge = null;
    let modeDescription = null;
    let themeInputValidation = null;

    // 文件夹选择相关DOM元素
    let selectFolderBtn = null;
    let folderUploadArea = null;
    let folderSelectionResult = null;
    let selectedFolderPath = null;
    let rscFileStatus = null;
    let ugcFileStatus = null;

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
        themeSelector = document.getElementById('themeSelector');
        processThemeBtn = document.getElementById('processThemeBtn');
        resetBtn = document.getElementById('resetBtn');
        enableDirectSaveBtn = document.getElementById('enableDirectSaveBtn');
        enableUGCDirectSaveBtn = document.getElementById('enableUGCDirectSaveBtn');
        fileStatus = document.getElementById('fileStatus');
        operationModeIndicator = document.getElementById('operationModeIndicator');
        modeBadge = document.getElementById('modeBadge');
        modeDescription = document.getElementById('modeDescription');
        themeInputValidation = document.getElementById('themeInputValidation');

        // 获取文件夹选择相关DOM元素
        selectFolderBtn = document.getElementById('selectFolderBtn');
        folderUploadArea = document.getElementById('folderUploadArea');
        folderSelectionResult = document.getElementById('folderSelectionResult');
        selectedFolderPath = document.getElementById('selectedFolderPath');
        rscFileStatus = document.getElementById('folderRscFileStatus');
        ugcFileStatus = document.getElementById('folderUgcFileStatus');

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

        // 初始化文件夹选择功能
        initializeFolderSelection();

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

        // 主题选择器事件监听
        if (themeSelector) {
            themeSelector.addEventListener('change', handleThemeSelection);
        }

        // 主题输入框事件监听
        if (themeNameInput) {
            themeNameInput.addEventListener('input', handleThemeInput);
            themeNameInput.addEventListener('blur', validateThemeInput);
        }

        // 初始化多语言功能
        initMultiLanguageFeatures();

        // 初始化UGC配置验证
        initUGCValidation();

        // 初始化Light配置验证
        initLightValidation();

        // 初始化ColorInfo配置验证
        initColorInfoValidation();
    }

    /**
     * 初始化多语言功能
     */
    function initMultiLanguageFeatures() {
        const openMultiLangBtn = document.getElementById('openMultiLangBtn');
        const multiLangIdInput = document.getElementById('multiLangId');

        if (openMultiLangBtn) {
            openMultiLangBtn.addEventListener('click', handleOpenMultiLangTable);
        }

        // 验证多语言ID输入
        if (multiLangIdInput) {
            multiLangIdInput.addEventListener('input', validateMultiLangId);
        }
    }

    /**
     * 初始化UGC配置验证
     */
    function initUGCValidation() {
        // 透明度字段验证
        const transparencyFields = [
            'fragileGlassAlpha', 'fragilePatternAlpha',
            'fragileActiveGlassAlpha', 'fragileActivePatternAlpha'
        ];

        transparencyFields.forEach(fieldId => {
            const input = document.getElementById(fieldId);
            if (input) {
                input.addEventListener('input', function() {
                    validateTransparencyInput(this);
                });
                input.addEventListener('blur', function() {
                    validateTransparencyInput(this);
                });
            }
        });

        console.log('UGC配置验证已初始化');
    }

    /**
     * 初始化Light配置验证
     */
    function initLightValidation() {
        // 明度偏移字段验证 (-255 到 255)
        const offsetFields = ['lightMax', 'lightDark', 'lightMin'];

        offsetFields.forEach(fieldId => {
            const input = document.getElementById(fieldId);
            if (input) {
                input.addEventListener('input', function() {
                    validateLightOffsetInput(this);
                });
                input.addEventListener('blur', function() {
                    validateLightOffsetInput(this);
                });
            }
        });

        // 高光等级字段验证 (0 到 1000)
        const specularLevelInput = document.getElementById('lightSpecularLevel');
        if (specularLevelInput) {
            specularLevelInput.addEventListener('input', function() {
                validateSpecularLevelInput(this);
            });
            specularLevelInput.addEventListener('blur', function() {
                validateSpecularLevelInput(this);
            });
        }

        // 光泽度字段验证 (10 到 1000)
        const glossInput = document.getElementById('lightGloss');
        if (glossInput) {
            glossInput.addEventListener('input', function() {
                validateGlossInput(this);
            });
            glossInput.addEventListener('blur', function() {
                validateGlossInput(this);
            });
        }

        // 颜色字段验证
        const colorInput = document.getElementById('lightSpecularColor');
        if (colorInput) {
            colorInput.addEventListener('input', function() {
                validateColorInput(this);
            });
            colorInput.addEventListener('blur', function() {
                validateColorInput(this);
            });
        }

        console.log('Light配置验证已初始化');
    }

    /**
     * 初始化ColorInfo配置验证
     */
    function initColorInfoValidation() {
        console.log('初始化ColorInfo配置验证...');

        // RGB字段验证 (0 到 255)
        const rgbFields = [
            'PickupDiffR', 'PickupDiffG', 'PickupDiffB',
            'PickupReflR', 'PickupReflG', 'PickupReflB',
            'BallSpecR', 'BallSpecG', 'BallSpecB',
            'ForegroundFogR', 'ForegroundFogG', 'ForegroundFogB'
        ];

        rgbFields.forEach(fieldId => {
            const input = document.getElementById(fieldId);
            if (input) {
                input.addEventListener('input', function() {
                    validateRgbInput(this);
                    updateRgbColorPreview(fieldId);
                });
                input.addEventListener('blur', function() {
                    validateRgbInput(this);
                    updateRgbColorPreview(fieldId);
                });
            }
        });

        // 雾开始距离验证 (0 到 40)
        const fogStartInput = document.getElementById('FogStart');
        if (fogStartInput) {
            fogStartInput.addEventListener('input', function() {
                validateFogStartInput(this);
            });
            fogStartInput.addEventListener('blur', function() {
                validateFogStartInput(this);
            });
        }

        // 雾结束距离验证 (0 到 90)
        const fogEndInput = document.getElementById('FogEnd');
        if (fogEndInput) {
            fogEndInput.addEventListener('input', function() {
                validateFogEndInput(this);
            });
            fogEndInput.addEventListener('blur', function() {
                validateFogEndInput(this);
            });
        }

        console.log('ColorInfo配置验证已初始化');
    }

    /**
     * 验证透明度输入框
     */
    function validateTransparencyInput(input) {
        const value = parseInt(input.value);
        const fieldName = input.id;

        // 移除之前的错误样式
        input.classList.remove('validation-error');

        // 移除之前的错误提示
        const existingError = input.parentElement.querySelector('.validation-message');
        if (existingError) {
            existingError.remove();
        }

        if (isNaN(value) || value < 0 || value > 100) {
            // 添加错误样式
            input.classList.add('validation-error');

            // 创建错误提示
            const errorMsg = document.createElement('div');
            errorMsg.className = 'validation-message error';
            errorMsg.textContent = '透明度值必须在0-100之间';
            input.parentElement.appendChild(errorMsg);

            // 自动修正值
            if (isNaN(value) || value < 0) {
                input.value = '0';
            } else if (value > 100) {
                input.value = '100';
            }

            console.warn(`${fieldName} 透明度值无效: ${input.value}，已自动修正`);
        } else {
            console.log(`${fieldName} 透明度值有效: ${value}`);
        }
    }

    /**
     * 验证明度偏移输入框 (-255 到 255)
     */
    function validateLightOffsetInput(input) {
        const value = parseInt(input.value);
        const fieldName = input.id;

        // 移除之前的错误样式
        input.classList.remove('validation-error');

        // 移除之前的错误提示
        const existingError = input.parentElement.querySelector('.validation-message');
        if (existingError) {
            existingError.remove();
        }

        if (isNaN(value) || value < -255 || value > 255) {
            // 添加错误样式
            input.classList.add('validation-error');

            // 创建错误提示
            const errorMsg = document.createElement('div');
            errorMsg.className = 'validation-message error';
            errorMsg.textContent = '明度偏移值必须在-255到255之间';
            input.parentElement.appendChild(errorMsg);

            // 自动修正值
            if (isNaN(value)) {
                input.value = '0';
            } else if (value < -255) {
                input.value = '-255';
            } else if (value > 255) {
                input.value = '255';
            }

            console.warn(`${fieldName} 明度偏移值无效: ${input.value}，已自动修正`);
        } else {
            console.log(`${fieldName} 明度偏移值有效: ${value}`);
        }
    }

    /**
     * 验证高光等级输入框 (0 到 1000)
     */
    function validateSpecularLevelInput(input) {
        const value = parseInt(input.value);
        const fieldName = input.id;

        // 移除之前的错误样式
        input.classList.remove('validation-error');

        // 移除之前的错误提示
        const existingError = input.parentElement.querySelector('.validation-message');
        if (existingError) {
            existingError.remove();
        }

        if (isNaN(value) || value < 0 || value > 1000) {
            // 添加错误样式
            input.classList.add('validation-error');

            // 创建错误提示
            const errorMsg = document.createElement('div');
            errorMsg.className = 'validation-message error';
            errorMsg.textContent = '高光等级值必须在0到1000之间';
            input.parentElement.appendChild(errorMsg);

            // 自动修正值
            if (isNaN(value) || value < 0) {
                input.value = '0';
            } else if (value > 1000) {
                input.value = '1000';
            }

            console.warn(`${fieldName} 高光等级值无效: ${input.value}，已自动修正`);
        } else {
            console.log(`${fieldName} 高光等级值有效: ${value}`);
        }
    }

    /**
     * 验证光泽度输入框 (10 到 1000)
     */
    function validateGlossInput(input) {
        const value = parseInt(input.value);
        const fieldName = input.id;

        // 移除之前的错误样式
        input.classList.remove('validation-error');

        // 移除之前的错误提示
        const existingError = input.parentElement.querySelector('.validation-message');
        if (existingError) {
            existingError.remove();
        }

        if (isNaN(value) || value < 10 || value > 1000) {
            // 添加错误样式
            input.classList.add('validation-error');

            // 创建错误提示
            const errorMsg = document.createElement('div');
            errorMsg.className = 'validation-message error';
            errorMsg.textContent = '光泽度值必须在10到1000之间';
            input.parentElement.appendChild(errorMsg);

            // 自动修正值
            if (isNaN(value) || value < 10) {
                input.value = '10';
            } else if (value > 1000) {
                input.value = '1000';
            }

            console.warn(`${fieldName} 光泽度值无效: ${input.value}，已自动修正`);
        } else {
            console.log(`${fieldName} 光泽度值有效: ${value}`);
        }
    }

    /**
     * 验证颜色输入框 (16进制格式)
     */
    function validateColorInput(input) {
        const value = input.value.trim().toUpperCase();
        const fieldName = input.id;

        // 移除之前的错误样式
        input.classList.remove('validation-error');

        // 移除之前的错误提示
        const existingError = input.parentElement.querySelector('.validation-message');
        if (existingError) {
            existingError.remove();
        }

        if (!value || !/^[0-9A-F]{6}$/i.test(value)) {
            // 添加错误样式
            input.classList.add('validation-error');

            // 创建错误提示
            const errorMsg = document.createElement('div');
            errorMsg.className = 'validation-message error';
            errorMsg.textContent = '请输入6位16进制颜色值 (如: FFFFFF)';
            input.parentElement.appendChild(errorMsg);

            // 自动修正值
            if (!value) {
                input.value = 'FFFFFF';
            } else {
                // 尝试修正常见错误
                const cleaned = value.replace(/[^0-9A-F]/gi, '');
                if (cleaned.length >= 6) {
                    input.value = cleaned.substring(0, 6);
                } else {
                    input.value = 'FFFFFF';
                }
            }

            console.warn(`${fieldName} 颜色值无效: ${value}，已自动修正为: ${input.value}`);
        } else {
            // 更新颜色预览
            if (window.App && window.App.ColorPicker) {
                window.App.ColorPicker.updateColorPreview(fieldName, value);
            }
            console.log(`${fieldName} 颜色值有效: ${value}`);
        }
    }

    /**
     * 验证RGB输入框 (0 到 255)
     */
    function validateRgbInput(input) {
        const value = parseInt(input.value);
        const fieldName = input.id;

        // 移除之前的错误样式
        input.classList.remove('validation-error');

        // 移除之前的错误提示
        const existingError = input.parentElement.querySelector('.validation-message');
        if (existingError) {
            existingError.remove();
        }

        if (isNaN(value) || value < 0 || value > 255) {
            // 添加错误样式
            input.classList.add('validation-error');

            // 创建错误提示
            const errorMsg = document.createElement('div');
            errorMsg.className = 'validation-message error';
            errorMsg.textContent = 'RGB值范围: 0-255';
            input.parentElement.appendChild(errorMsg);

            // 自动修正值
            const correctedValue = Math.max(0, Math.min(255, isNaN(value) ? 255 : value));
            input.value = correctedValue;

            console.warn(`${fieldName} RGB值无效: ${input.value}，已自动修正为: ${correctedValue}`);
        } else {
            console.log(`${fieldName} RGB值有效: ${value}`);
        }
    }

    /**
     * 验证雾开始距离输入框 (0 到 40)
     */
    function validateFogStartInput(input) {
        const value = parseInt(input.value);
        const fieldName = input.id;

        // 移除之前的错误样式
        input.classList.remove('validation-error');

        // 移除之前的错误提示
        const existingError = input.parentElement.querySelector('.validation-message');
        if (existingError) {
            existingError.remove();
        }

        if (isNaN(value) || value < 0 || value > 40) {
            // 添加错误样式
            input.classList.add('validation-error');

            // 创建错误提示
            const errorMsg = document.createElement('div');
            errorMsg.className = 'validation-message error';
            errorMsg.textContent = '雾开始距离范围: 0-40';
            input.parentElement.appendChild(errorMsg);

            // 自动修正值
            const correctedValue = Math.max(0, Math.min(40, isNaN(value) ? 10 : value));
            input.value = correctedValue;

            console.warn(`${fieldName} 雾开始距离无效: ${input.value}，已自动修正为: ${correctedValue}`);
        } else {
            console.log(`${fieldName} 雾开始距离有效: ${value}`);
        }
    }

    /**
     * 验证雾结束距离输入框 (0 到 90)
     */
    function validateFogEndInput(input) {
        const value = parseInt(input.value);
        const fieldName = input.id;

        // 移除之前的错误样式
        input.classList.remove('validation-error');

        // 移除之前的错误提示
        const existingError = input.parentElement.querySelector('.validation-message');
        if (existingError) {
            existingError.remove();
        }

        if (isNaN(value) || value < 0 || value > 90) {
            // 添加错误样式
            input.classList.add('validation-error');

            // 创建错误提示
            const errorMsg = document.createElement('div');
            errorMsg.className = 'validation-message error';
            errorMsg.textContent = '雾结束距离范围: 0-90';
            input.parentElement.appendChild(errorMsg);

            // 自动修正值
            const correctedValue = Math.max(0, Math.min(90, isNaN(value) ? 50 : value));
            input.value = correctedValue;

            console.warn(`${fieldName} 雾结束距离无效: ${input.value}，已自动修正为: ${correctedValue}`);
        } else {
            console.log(`${fieldName} 雾结束距离有效: ${value}`);
        }
    }

    /**
     * 更新RGB颜色预览
     */
    function updateRgbColorPreview(fieldId) {
        // 确定RGB组
        let rgbGroup = '';
        if (fieldId.startsWith('PickupDiff')) {
            rgbGroup = 'PickupDiff';
        } else if (fieldId.startsWith('PickupRefl')) {
            rgbGroup = 'PickupRefl';
        } else if (fieldId.startsWith('BallSpec')) {
            rgbGroup = 'BallSpec';
        } else if (fieldId.startsWith('ForegroundFog')) {
            rgbGroup = 'ForegroundFog';
        } else {
            return;
        }

        // 获取RGB值
        const rInput = document.getElementById(rgbGroup + 'R');
        const gInput = document.getElementById(rgbGroup + 'G');
        const bInput = document.getElementById(rgbGroup + 'B');
        const preview = document.getElementById(rgbGroup + 'ColorPreview');

        if (rInput && gInput && bInput && preview) {
            const r = parseInt(rInput.value) || 0;
            const g = parseInt(gInput.value) || 0;
            const b = parseInt(bInput.value) || 0;

            preview.style.backgroundColor = `rgb(${r}, ${g}, ${b})`;
        }
    }

    /**
     * 处理打开多语言表
     */
    function handleOpenMultiLangTable() {
        // 获取当前主题名称
        const themeName = document.getElementById('themeNameInput')?.value.trim();

        if (!themeName) {
            App.Utils.showStatus('请先输入新主题名称', 'warning');
            document.getElementById('themeNameInput')?.focus();
            return;
        }

        // 显示确认对话框，避免弹窗拦截
        const confirmMessage = `🔗 即将打开外部链接

将要打开在线多语言表进行配置：
https://www.kdocs.cn/l/cuwWQPWT7HPY

请在表格中添加以下信息：
• 主题名称：${themeName}
• 记录系统分配的多语言ID

确认打开外部链接吗？`;

        if (confirm(confirmMessage)) {
            try {
                // 打开在线多语言表
                const newWindow = window.open('https://www.kdocs.cn/l/cuwWQPWT7HPY', '_blank');

                if (newWindow) {
                    // 成功打开
                    App.Utils.showStatus('多语言表已打开，请填写完成后回来输入多语言ID', 'info', 5000);

                    // 高亮多语言ID输入框
                    setTimeout(() => {
                        const multiLangIdInput = document.getElementById('multiLangId');
                        if (multiLangIdInput) {
                            multiLangIdInput.focus();
                            multiLangIdInput.style.border = '2px solid #ffc107';
                            multiLangIdInput.placeholder = '请输入在线表中分配的多语言ID';
                        }
                    }, 1000);
                } else {
                    // 弹窗被拦截
                    App.Utils.showStatus('弹窗被浏览器拦截，请手动打开链接：https://www.kdocs.cn/l/cuwWQPWT7HPY', 'warning', 8000);
                }
            } catch (error) {
                console.error('打开多语言表失败:', error);
                App.Utils.showStatus('无法打开多语言表，请手动访问：https://www.kdocs.cn/l/cuwWQPWT7HPY', 'error', 8000);
            }
        }
    }

    /**
     * 验证多语言ID
     */
    function validateMultiLangId() {
        const multiLangIdInput = document.getElementById('multiLangId');
        const value = multiLangIdInput.value;

        if (value && !isNaN(value) && parseInt(value) > 0) {
            multiLangIdInput.style.border = '2px solid #4CAF50';
            App.Utils.showStatus('多语言ID已设置', 'success', 2000);
            updateMultiLangConfig();
        } else if (value) {
            multiLangIdInput.style.border = '2px solid #f44336';
        } else {
            multiLangIdInput.style.border = '1px solid rgba(255, 255, 255, 0.3)';
        }
    }

    /**
     * 更新多语言配置
     */
    function updateMultiLangConfig() {
        const themeName = document.getElementById('themeNameInput')?.value.trim();
        const multiLangId = document.getElementById('multiLangId')?.value.trim();

        multiLangConfig = {
            displayName: themeName || '',
            id: multiLangId ? parseInt(multiLangId) : null,
            isValid: themeName && multiLangId && !isNaN(multiLangId) && parseInt(multiLangId) > 0
        };

        console.log('多语言配置已更新:', multiLangConfig);
    }

    /**
     * 获取用户输入的多语言配置
     */
    function getMultiLanguageConfig() {
        updateMultiLangConfig();
        return multiLangConfig || {
            displayName: '',
            id: null,
            isValid: false
        };
    }

    /**
     * 显示/隐藏多语言配置面板
     */
    function toggleMultiLangPanel(show) {
        const panel = document.getElementById('multiLangConfigSection');
        if (panel) {
            panel.style.display = show ? 'block' : 'none';

            // 如果隐藏面板，清空输入
            if (!show) {
                const multiLangIdInput = document.getElementById('multiLangId');
                if (multiLangIdInput) {
                    multiLangIdInput.value = '';
                    multiLangIdInput.style.border = '1px solid rgba(255, 255, 255, 0.3)';
                    multiLangIdInput.placeholder = '请先打开在线表填写主题信息，然后输入分配的多语言ID';
                }
                multiLangConfig = null;
            }
        }
    }

    /**
     * 显示/隐藏UGC配置面板
     */
    function toggleUGCConfigPanel(show) {
        const panel = document.getElementById('ugcConfigSection');
        if (panel) {
            panel.style.display = show ? 'block' : 'none';

            // 如果显示面板，重新绑定图案选择器事件
            if (show && window.App && window.App.UGCPatternSelector) {
                setTimeout(() => {
                    if (typeof window.App.UGCPatternSelector.rebindButtonEvents === 'function') {
                        window.App.UGCPatternSelector.rebindButtonEvents();
                    }
                }, 100); // 延迟一点确保DOM更新完成
            }

            // 如果隐藏面板，重置为默认值
            if (!show) {
                resetUGCConfigToDefaults();
            }

            console.log('UGC配置面板', show ? '已显示' : '已隐藏');
        }
    }

    /**
     * 显示/隐藏Light配置面板
     */
    function toggleLightConfigPanel(show) {
        const panel = document.getElementById('lightConfigSection');
        if (panel) {
            panel.style.display = show ? 'block' : 'none';
            console.log('Light配置面板', show ? '已显示' : '已隐藏');
        }
    }

    /**
     * 显示/隐藏ColorInfo配置面板
     */
    function toggleColorInfoConfigPanel(show) {
        const panel = document.getElementById('colorinfoConfigSection');
        if (panel) {
            panel.style.display = show ? 'block' : 'none';
            console.log('ColorInfo配置面板', show ? '已显示' : '已隐藏');
        }
    }

    /**
     * 重置UGC配置为默认值
     */
    function resetUGCConfigToDefaults() {
        const ugcFields = [
            'groundPatternIndex', 'groundFrameIndex',
            'fragilePatternIndex', 'fragileFrameIndex', 'fragileGlassAlpha', 'fragilePatternAlpha',
            'fragileActivePatternIndex', 'fragileActiveFrameIndex', 'fragileActiveGlassAlpha', 'fragileActivePatternAlpha',
            'jumpPatternIndex', 'jumpFrameIndex',
            'jumpActivePatternIndex', 'jumpActiveFrameIndex'
        ];

        ugcFields.forEach(fieldId => {
            const input = document.getElementById(fieldId);
            if (input) {
                if (fieldId.includes('Alpha')) {
                    input.value = '50'; // 透明度默认50
                } else {
                    input.value = '0'; // 图案ID默认0
                }
            }
        });
    }

    /**
     * 获取表中最后一个主题的Light配置数据
     */
    function getLastThemeLightConfig() {
        if (!rscAllSheetsData || !rscAllSheetsData['Light']) {
            console.log('RSC_Theme Light数据未加载，使用硬编码默认值');
            return {
                lightMax: '0',
                lightDark: '0',
                lightMin: '0',
                lightSpecularLevel: '100',
                lightGloss: '100',
                lightSpecularColor: 'FFFFFF'
            };
        }

        const lightData = rscAllSheetsData['Light'];
        const lightHeaderRow = lightData[0];
        const lightNotesColumnIndex = lightHeaderRow.findIndex(col => col === 'notes');

        if (lightNotesColumnIndex === -1 || lightData.length <= 1) {
            console.log('RSC_Theme Light sheet没有notes列或没有数据，使用硬编码默认值');
            return {
                lightMax: '0',
                lightDark: '0',
                lightMin: '0',
                lightSpecularLevel: '100',
                lightGloss: '100',
                lightSpecularColor: 'FFFFFF'
            };
        }

        // 获取最后一行数据（跳过标题行）
        const lastRowIndex = lightData.length - 1;
        const lastRow = lightData[lastRowIndex];

        console.log(`读取表中最后一个主题的Light配置，行索引: ${lastRowIndex}`);

        // 构建字段映射
        const lightFieldMapping = {
            'Max': 'lightMax',
            'Dark': 'lightDark',
            'Min': 'lightMin',
            'SpecularLevel': 'lightSpecularLevel',
            'Gloss': 'lightGloss',
            'SpecularColor': 'lightSpecularColor'
        };

        const lastThemeConfig = {};
        Object.entries(lightFieldMapping).forEach(([columnName, fieldId]) => {
            const columnIndex = lightHeaderRow.findIndex(col => col === columnName);
            if (columnIndex !== -1) {
                const value = lastRow[columnIndex];
                lastThemeConfig[fieldId] = (value !== undefined && value !== null && value !== '') ? value.toString() : '0';
            } else {
                // 如果找不到列，使用默认值
                const defaults = {
                    'lightMax': '0',
                    'lightDark': '0',
                    'lightMin': '0',
                    'lightSpecularLevel': '100',
                    'lightGloss': '100',
                    'lightSpecularColor': 'FFFFFF'
                };
                lastThemeConfig[fieldId] = defaults[fieldId] || '0';
            }
        });

        console.log('最后一个主题的Light配置:', lastThemeConfig);
        return lastThemeConfig;
    }

    /**
     * 重置Light配置为默认值（新建主题时使用表中最后一个主题的数据）
     */
    function resetLightConfigToDefaults() {
        const lightDefaults = getLastThemeLightConfig();

        Object.entries(lightDefaults).forEach(([fieldId, defaultValue]) => {
            const input = document.getElementById(fieldId);
            if (input) {
                input.value = defaultValue;

                // 更新颜色预览
                if (fieldId === 'lightSpecularColor') {
                    updateColorPreview(fieldId, defaultValue);
                }
            }
        });

        console.log('Light配置已重置为最后一个主题的配置');
    }

    /**
     * 获取表中最后一个主题的ColorInfo配置数据
     */
    function getLastThemeColorInfoConfig() {
        if (!rscAllSheetsData || !rscAllSheetsData['ColorInfo']) {
            console.log('RSC_Theme ColorInfo数据未加载，使用硬编码默认值');
            return {
                PickupDiffR: '255',
                PickupDiffG: '255',
                PickupDiffB: '255',
                PickupReflR: '255',
                PickupReflG: '255',
                PickupReflB: '255',
                BallSpecR: '255',
                BallSpecG: '255',
                BallSpecB: '255',
                ForegroundFogR: '128',
                ForegroundFogG: '128',
                ForegroundFogB: '128',
                FogStart: '10',
                FogEnd: '50'
            };
        }

        const colorInfoData = rscAllSheetsData['ColorInfo'];
        const colorInfoHeaderRow = colorInfoData[0];
        const colorInfoNotesColumnIndex = colorInfoHeaderRow.findIndex(col => col === 'notes');

        if (colorInfoNotesColumnIndex === -1 || colorInfoData.length <= 1) {
            console.log('RSC_Theme ColorInfo sheet没有notes列或没有数据，使用硬编码默认值');
            return {
                PickupDiffR: '255',
                PickupDiffG: '255',
                PickupDiffB: '255',
                PickupReflR: '255',
                PickupReflG: '255',
                PickupReflB: '255',
                ForegroundFogR: '128',
                ForegroundFogG: '128',
                ForegroundFogB: '128',
                FogStart: '10',
                FogEnd: '50'
            };
        }

        // 获取最后一行数据（跳过标题行）
        const lastRowIndex = colorInfoData.length - 1;
        const lastRow = colorInfoData[lastRowIndex];

        console.log(`读取表中最后一个主题的ColorInfo配置，行索引: ${lastRowIndex}`);

        // 构建字段映射
        const colorInfoFieldMapping = {
            'PickupDiffR': 'PickupDiffR',
            'PickupDiffG': 'PickupDiffG',
            'PickupDiffB': 'PickupDiffB',
            'PickupReflR': 'PickupReflR',
            'PickupReflG': 'PickupReflG',
            'PickupReflB': 'PickupReflB',
            'BallSpecR': 'BallSpecR',
            'BallSpecG': 'BallSpecG',
            'BallSpecB': 'BallSpecB',
            'ForegroundFogR': 'ForegroundFogR',
            'ForegroundFogG': 'ForegroundFogG',
            'ForegroundFogB': 'ForegroundFogB',
            'FogStart': 'FogStart',
            'FogEnd': 'FogEnd'
        };

        const lastThemeConfig = {};
        Object.entries(colorInfoFieldMapping).forEach(([columnName, fieldId]) => {
            const columnIndex = colorInfoHeaderRow.findIndex(col => col === columnName);
            if (columnIndex !== -1) {
                const value = lastRow[columnIndex];
                lastThemeConfig[fieldId] = (value !== undefined && value !== null && value !== '') ? value.toString() : '0';
            } else {
                // 如果找不到列，使用默认值
                const defaults = {
                    'PickupDiffR': '255',
                    'PickupDiffG': '255',
                    'PickupDiffB': '255',
                    'PickupReflR': '255',
                    'PickupReflG': '255',
                    'PickupReflB': '255',
                    'BallSpecR': '255',
                    'BallSpecG': '255',
                    'BallSpecB': '255',
                    'ForegroundFogR': '128',
                    'ForegroundFogG': '128',
                    'ForegroundFogB': '128',
                    'FogStart': '10',
                    'FogEnd': '50'
                };
                lastThemeConfig[fieldId] = defaults[fieldId] || '0';
            }
        });

        console.log('最后一个主题的ColorInfo配置:', lastThemeConfig);
        return lastThemeConfig;
    }

    /**
     * 重置ColorInfo配置为默认值（新建主题时使用表中最后一个主题的数据）
     */
    function resetColorInfoConfigToDefaults() {
        const colorInfoDefaults = getLastThemeColorInfoConfig();

        Object.entries(colorInfoDefaults).forEach(([fieldId, defaultValue]) => {
            const input = document.getElementById(fieldId);
            if (input) {
                input.value = defaultValue;
                input.classList.remove('validation-error');

                // 移除错误提示
                const errorMsg = input.parentElement.querySelector('.validation-message');
                if (errorMsg) {
                    errorMsg.remove();
                }
            }
        });

        // 更新颜色预览
        updateRgbColorPreview('PickupDiffR');
        updateRgbColorPreview('PickupReflR');
        updateRgbColorPreview('BallSpecR');
        updateRgbColorPreview('ForegroundFogR');

        console.log('ColorInfo配置已重置为最后一个主题的配置');
    }

    /**
     * 验证透明度值（0-100范围）
     */
    function validateTransparency(value, defaultValue = 50) {
        const numValue = parseInt(value);
        if (isNaN(numValue) || numValue < 0 || numValue > 100) {
            return defaultValue;
        }
        return numValue;
    }

    /**
     * 验证边框ID值（只能是0或1）
     */
    function validateFrameIndex(value, defaultValue = 0) {
        const numValue = parseInt(value);
        if (numValue !== 0 && numValue !== 1) {
            return defaultValue;
        }
        return numValue;
    }



    /**
     * 验证明度偏移值（-255到255范围）
     */
    function validateLightOffset(value, defaultValue = 0) {
        const numValue = parseInt(value);
        if (isNaN(numValue) || numValue < -255 || numValue > 255) {
            return defaultValue;
        }
        return numValue;
    }

    /**
     * 验证高光等级值（0到1000范围）
     */
    function validateSpecularValue(value, defaultValue = 100) {
        const numValue = parseInt(value);
        if (isNaN(numValue) || numValue < 0 || numValue > 1000) {
            return defaultValue;
        }
        return numValue;
    }

    /**
     * 验证光泽度值（10到1000范围）
     */
    function validateGlossValue(value, defaultValue = 100) {
        const numValue = parseInt(value);
        if (isNaN(numValue) || numValue < 10 || numValue > 1000) {
            return defaultValue;
        }
        return numValue;
    }

    /**
     * 验证16进制颜色值
     */
    function validateHexColor(value, defaultValue = 'FFFFFF') {
        if (!value || typeof value !== 'string') {
            return defaultValue;
        }

        const cleanValue = value.replace('#', '').toUpperCase();
        if (/^[0-9A-F]{6}$/i.test(cleanValue)) {
            return cleanValue;
        }

        return defaultValue;
    }

    /**
     * 验证RGB值（0-255范围）
     */
    function validateRgbValue(value, defaultValue = 255) {
        const numValue = parseInt(value);
        if (isNaN(numValue) || numValue < 0 || numValue > 255) {
            return defaultValue;
        }
        return numValue;
    }

    /**
     * 验证雾开始距离（0-40范围）
     */
    function validateFogStart(value, defaultValue = 10) {
        const numValue = parseInt(value);
        if (isNaN(numValue) || numValue < 0 || numValue > 40) {
            return defaultValue;
        }
        return numValue;
    }

    /**
     * 验证雾结束距离（0-90范围）
     */
    function validateFogEnd(value, defaultValue = 50) {
        const numValue = parseInt(value);
        if (isNaN(numValue) || numValue < 0 || numValue > 90) {
            return defaultValue;
        }
        return numValue;
    }

    /**
     * 获取UGC配置数据
     */
    function getUGCConfigData() {
        return {
            groundPatternIndex: parseInt(document.getElementById('groundPatternIndex')?.value || '0'),
            groundFrameIndex: parseInt(document.getElementById('groundFrameIndex')?.value || '0'),
            fragilePatternIndex: parseInt(document.getElementById('fragilePatternIndex')?.value || '0'),
            fragileFrameIndex: validateFrameIndex(document.getElementById('fragileFrameIndex')?.value, 0),
            fragileGlassAlpha: validateTransparency(document.getElementById('fragileGlassAlpha')?.value, 50),
            fragilePatternAlpha: validateTransparency(document.getElementById('fragilePatternAlpha')?.value, 50),
            fragileActivePatternIndex: parseInt(document.getElementById('fragileActivePatternIndex')?.value || '0'),
            fragileActiveFrameIndex: validateFrameIndex(document.getElementById('fragileActiveFrameIndex')?.value, 0),
            fragileActiveGlassAlpha: validateTransparency(document.getElementById('fragileActiveGlassAlpha')?.value, 50),
            fragileActivePatternAlpha: validateTransparency(document.getElementById('fragileActivePatternAlpha')?.value, 50),
            jumpPatternIndex: parseInt(document.getElementById('jumpPatternIndex')?.value || '0'),
            jumpFrameIndex: parseInt(document.getElementById('jumpFrameIndex')?.value || '0'),
            jumpActivePatternIndex: parseInt(document.getElementById('jumpActivePatternIndex')?.value || '0'),
            jumpActiveFrameIndex: parseInt(document.getElementById('jumpActiveFrameIndex')?.value || '0')
        };
    }

    /**
     * 获取Light配置数据
     */
    function getLightConfigData() {
        return {
            Max: validateLightOffset(document.getElementById('lightMax')?.value, 0),
            Dark: validateLightOffset(document.getElementById('lightDark')?.value, 0),
            Min: validateLightOffset(document.getElementById('lightMin')?.value, 0),
            SpecularLevel: validateSpecularValue(document.getElementById('lightSpecularLevel')?.value, 100),
            Gloss: validateGlossValue(document.getElementById('lightGloss')?.value, 100),
            SpecularColor: validateHexColor(document.getElementById('lightSpecularColor')?.value, 'FFFFFF')
        };
    }

    /**
     * 获取ColorInfo配置数据
     */
    function getColorInfoConfigData() {
        return {
            PickupDiffR: validateRgbValue(document.getElementById('PickupDiffR')?.value, 255),
            PickupDiffG: validateRgbValue(document.getElementById('PickupDiffG')?.value, 255),
            PickupDiffB: validateRgbValue(document.getElementById('PickupDiffB')?.value, 255),
            PickupReflR: validateRgbValue(document.getElementById('PickupReflR')?.value, 255),
            PickupReflG: validateRgbValue(document.getElementById('PickupReflG')?.value, 255),
            PickupReflB: validateRgbValue(document.getElementById('PickupReflB')?.value, 255),
            BallSpecR: validateRgbValue(document.getElementById('BallSpecR')?.value, 255),
            BallSpecG: validateRgbValue(document.getElementById('BallSpecG')?.value, 255),
            BallSpecB: validateRgbValue(document.getElementById('BallSpecB')?.value, 255),
            ForegroundFogR: validateRgbValue(document.getElementById('ForegroundFogR')?.value, 128),
            ForegroundFogG: validateRgbValue(document.getElementById('ForegroundFogG')?.value, 128),
            ForegroundFogB: validateRgbValue(document.getElementById('ForegroundFogB')?.value, 128),
            FogStart: validateFogStart(document.getElementById('FogStart')?.value, 10),
            FogEnd: validateFogEnd(document.getElementById('FogEnd')?.value, 50)
        };
    }

    /**
     * 设置UGC配置数据（用于更新现有主题时显示当前值）
     */
    function setUGCConfigData(configData) {
        if (!configData) return;

        const fieldMapping = {
            'groundPatternIndex': configData.groundPatternIndex,
            'groundFrameIndex': configData.groundFrameIndex,
            'fragilePatternIndex': configData.fragilePatternIndex,
            'fragileFrameIndex': configData.fragileFrameIndex,
            'fragileGlassAlpha': configData.fragileGlassAlpha,
            'fragilePatternAlpha': configData.fragilePatternAlpha,
            'fragileActivePatternIndex': configData.fragileActivePatternIndex,
            'fragileActiveFrameIndex': configData.fragileActiveFrameIndex,
            'fragileActiveGlassAlpha': configData.fragileActiveGlassAlpha,
            'fragileActivePatternAlpha': configData.fragileActivePatternAlpha,
            'jumpPatternIndex': configData.jumpPatternIndex,
            'jumpFrameIndex': configData.jumpFrameIndex,
            'jumpActivePatternIndex': configData.jumpActivePatternIndex,
            'jumpActiveFrameIndex': configData.jumpActiveFrameIndex
        };

        Object.entries(fieldMapping).forEach(([fieldId, value]) => {
            const input = document.getElementById(fieldId);
            if (input && value !== undefined) {
                input.value = value.toString();
            }
        });
    }

    /**
     * 加载现有主题的UGC配置
     */
    function loadExistingUGCConfig(themeName) {
        console.log('=== 开始加载现有UGC配置 ===');
        console.log('主题名称:', themeName);
        console.log('ugcAllSheetsData状态:', ugcAllSheetsData ? '已加载' : '未加载');
        console.log('rscAllSheetsData状态:', rscAllSheetsData ? '已加载' : '未加载');

        if (ugcAllSheetsData) {
            console.log('UGC数据包含的sheets:', Object.keys(ugcAllSheetsData));
        }

        if (!ugcAllSheetsData || !rscAllSheetsData || !themeName) {
            console.log('UGC数据或RSC数据未加载，或主题名称为空，使用默认值');
            resetUGCConfigToDefaults();
            return;
        }

        // 第一步：在RSC_Theme的Color表中找到主题对应的行号
        const rscColorData = rscAllSheetsData['Color'];
        if (!rscColorData || rscColorData.length === 0) {
            console.log('RSC_Theme的Color表未找到或为空，使用默认值');
            resetUGCConfigToDefaults();
            return;
        }

        const rscHeaderRow = rscColorData[0];
        const rscNotesColumnIndex = rscHeaderRow.findIndex(col => col === 'notes');

        if (rscNotesColumnIndex === -1) {
            console.log('RSC_Theme的Color表没有notes列，使用默认值');
            resetUGCConfigToDefaults();
            return;
        }

        // 查找主题在RSC中的行号
        const rscThemeRowIndex = rscColorData.findIndex((row, index) =>
            index > 0 && row[rscNotesColumnIndex] === themeName
        );

        if (rscThemeRowIndex === -1) {
            console.log(`在RSC_Theme的Color表中未找到主题 "${themeName}"，使用默认值`);
            resetUGCConfigToDefaults();
            return;
        }

        console.log(`在RSC_Theme的Color表中找到主题 "${themeName}"，行索引: ${rscThemeRowIndex}`);

        // 计算对应的数据行号（从1开始，因为第0行是表头）
        const targetRowNumber = rscThemeRowIndex;

        try {
            // 定义需要查找的sheet和对应的字段映射
            const sheetFieldMapping = {
                'Custom_Ground_Color': {
                    patternField: 'groundPatternIndex',
                    frameField: 'groundFrameIndex',
                    patternColumn: '_PatternUpIndex',
                    frameColumn: '_FrameIndex'
                },
                'Custom_Fragile_Color': {
                    patternField: 'fragilePatternIndex',
                    frameField: 'fragileFrameIndex',
                    glassAlphaField: 'fragileGlassAlpha',
                    patternAlphaField: 'fragilePatternAlpha',
                    patternColumn: '_PatternUpIndex',
                    frameColumn: '_FrameIndex',
                    glassAlphaColumn: '_GlassAlpha',
                    patternAlphaColumn: '_PatternAlpha'
                },
                'Custom_Fragile_Active_Color': {
                    patternField: 'fragileActivePatternIndex',
                    frameField: 'fragileActiveFrameIndex',
                    glassAlphaField: 'fragileActiveGlassAlpha',
                    patternAlphaField: 'fragileActivePatternAlpha',
                    patternColumn: '_PatternUpIndex',
                    frameColumn: '_FrameIndex',
                    glassAlphaColumn: '_GlassAlpha',
                    patternAlphaColumn: '_PatternAlpha'
                },
                'Custom_Jump_Color': {
                    patternField: 'jumpPatternIndex',
                    frameField: 'jumpFrameIndex',
                    patternColumn: '_PatternUpIndex',
                    frameColumn: '_FrameIndex'
                },
                'Custom_Jump_Active_Color': {
                    patternField: 'jumpActivePatternIndex',
                    frameField: 'jumpActiveFrameIndex',
                    patternColumn: '_PatternUpIndex',
                    frameColumn: '_FrameIndex'
                }
            };

            const configData = {};

            // 遍历每个sheet查找主题数据（使用行号匹配）
            Object.entries(sheetFieldMapping).forEach(([sheetName, mapping]) => {
                console.log(`\n--- 处理Sheet: ${sheetName} ---`);
                const sheetData = ugcAllSheetsData[sheetName];
                if (!sheetData || sheetData.length === 0) {
                    console.log(`Sheet ${sheetName} 不存在或为空`);
                    return;
                }

                console.log(`Sheet ${sheetName} 数据行数: ${sheetData.length}`);
                console.log(`目标行号: ${targetRowNumber}`);

                // 检查目标行是否存在
                if (targetRowNumber >= sheetData.length) {
                    console.log(`Sheet ${sheetName} 中不存在行号 ${targetRowNumber}（总行数: ${sheetData.length}）`);
                    return;
                }

                const headerRow = sheetData[0];
                const themeRow = sheetData[targetRowNumber];

                console.log(`Sheet ${sheetName} 表头:`, headerRow);
                console.log(`Sheet ${sheetName} 目标行数据:`, themeRow);

                // 提取字段值
                Object.entries(mapping).forEach(([fieldKey, fieldValue]) => {
                    if (fieldKey.endsWith('Field')) {
                        const columnName = mapping[fieldKey.replace('Field', 'Column')];
                        const columnIndex = headerRow.findIndex(col => col === columnName);

                        if (columnIndex !== -1) {
                            const value = themeRow[columnIndex];
                            configData[fieldValue] = value !== undefined && value !== '' ? parseInt(value) || 0 : 0;
                            console.log(`Sheet ${sheetName} 提取字段 ${columnName}(索引${columnIndex}) = ${value} -> ${fieldValue}`);
                        } else {
                            console.log(`Sheet ${sheetName} 未找到列 ${columnName}`);
                        }
                    }
                });
            });

            console.log('加载的UGC配置数据:', configData);
            setUGCConfigData(configData);

        } catch (error) {
            console.error('加载现有UGC配置失败:', error);
            resetUGCConfigToDefaults();
        }
    }

    /**
     * 更新主题类型指示器
     */
    function updateThemeTypeIndicator(smartConfig) {
        // 查找或创建主题类型提示元素
        let indicator = document.getElementById('themeTypeIndicator');
        if (!indicator) {
            indicator = document.createElement('div');
            indicator.id = 'themeTypeIndicator';
            indicator.className = 'theme-type-indicator';

            // 插入到主题输入框后面
            const themeInputGroup = themeNameInput?.parentElement;
            if (themeInputGroup) {
                themeInputGroup.appendChild(indicator);
            }
        }

        if (smartConfig.similarity.isSimilar) {
            // 同系列主题
            indicator.innerHTML = `
                <div class="indicator-content similar-theme">
                    <span class="indicator-icon">🔄</span>
                    <span class="indicator-text">检测到同系列主题，将自动复用 "${smartConfig.similarity.matchedTheme}" 的多语言配置</span>
                </div>
            `;
            indicator.className = 'theme-type-indicator similar-theme';
        } else {
            // 全新主题系列
            indicator.innerHTML = `
                <div class="indicator-content new-theme">
                    <span class="indicator-icon">✨</span>
                    <span class="indicator-text">检测到全新主题系列，需要配置多语言信息</span>
                </div>
            `;
            indicator.className = 'theme-type-indicator new-theme';
        }

        indicator.style.display = 'block';
    }

    /**
     * 清除主题类型指示器
     */
    function clearThemeTypeIndicator() {
        const indicator = document.getElementById('themeTypeIndicator');
        if (indicator) {
            indicator.style.display = 'none';
        }
    }

    /**
     * 提取主题名称的基础部分（去除数字后缀等）
     */
    function extractThemeBaseName(themeName) {
        if (!themeName) return '';

        // 去除首尾空格
        let baseName = themeName.trim();

        // 去除末尾的数字和常见分隔符
        // 匹配模式：数字、中文数字、罗马数字等
        baseName = baseName.replace(/[\s\-_]*[\d一二三四五六七八九十]+[\s\-_]*$/g, '');
        baseName = baseName.replace(/[\s\-_]*[IVXivx]+[\s\-_]*$/g, ''); // 罗马数字
        baseName = baseName.replace(/[\s\-_]*[第]*[\d一二三四五六七八九十]+[期代版]*[\s\-_]*$/g, ''); // 中文数字表达

        // 去除常见的版本标识
        baseName = baseName.replace(/[\s\-_]*(v|ver|version)[\d\.]*[\s\-_]*$/gi, '');
        baseName = baseName.replace(/[\s\-_]*(新|old|旧|原版|升级版|加强版)[\s\-_]*$/g, '');

        // 去除末尾的标点符号和空格
        baseName = baseName.replace(/[\s\-_\.\,\!\?\:]+$/, '');

        return baseName.trim();
    }

    /**
     * 检测主题名称相似性
     */
    function detectThemeSimilarity(newThemeName, existingThemes) {
        if (!newThemeName || !existingThemes || existingThemes.length === 0) {
            return { isSimilar: false, baseName: '', matchedTheme: null };
        }

        const newBaseName = extractThemeBaseName(newThemeName);
        console.log(`检测主题相似性: "${newThemeName}" -> 基础名称: "${newBaseName}"`);

        if (!newBaseName) {
            return { isSimilar: false, baseName: '', matchedTheme: null };
        }

        // 查找相似的现有主题
        for (const existingTheme of existingThemes) {
            const existingBaseName = extractThemeBaseName(existingTheme);
            console.log(`对比现有主题: "${existingTheme}" -> 基础名称: "${existingBaseName}"`);

            if (existingBaseName && newBaseName === existingBaseName) {
                console.log(`✅ 发现同系列主题: "${newThemeName}" 与 "${existingTheme}" 属于同系列`);
                return {
                    isSimilar: true,
                    baseName: newBaseName,
                    matchedTheme: existingTheme,
                    matchedBaseName: existingBaseName
                };
            }
        }

        console.log(`❌ 未发现相似主题，"${newThemeName}" 是全新主题系列`);
        return { isSimilar: false, baseName: newBaseName, matchedTheme: null };
    }

    /**
     * 获取现有主题列表
     */
    function getExistingThemeNames() {
        const themes = [];

        // 从主题选择器获取现有主题
        if (themeSelector) {
            const options = themeSelector.querySelectorAll('option');
            options.forEach(option => {
                if (option.value && option.value !== '') {
                    themes.push(option.value);
                }
            });
        }

        console.log('现有主题列表:', themes);
        return themes;
    }

    /**
     * 智能检测主题类型并获取多语言配置
     */
    function getSmartMultiLanguageConfig(themeName) {
        const existingThemes = getExistingThemeNames();
        const similarity = detectThemeSimilarity(themeName, existingThemes);

        // 获取用户手动输入的多语言配置
        const manualConfig = getMultiLanguageConfig();

        const result = {
            themeName: themeName,
            similarity: similarity,
            isNewSeries: !similarity.isSimilar,
            shouldShowConfig: !similarity.isSimilar, // 只有全新系列才显示配置面板
            multiLangConfig: null,
            source: 'none'
        };

        if (similarity.isSimilar) {
            // 同系列主题，尝试复用现有配置
            result.multiLangConfig = {
                displayName: themeName,
                id: null, // 将在UGCTheme处理时从现有主题获取
                isValid: true,
                isAutoDetected: true,
                basedOnTheme: similarity.matchedTheme
            };
            result.source = 'auto_detected';
            console.log(`🔄 同系列主题检测: 将复用 "${similarity.matchedTheme}" 的多语言配置`);
        } else if (manualConfig && manualConfig.isValid) {
            // 全新系列，使用用户手动配置
            result.multiLangConfig = manualConfig;
            result.source = 'manual_input';
            console.log(`✏️ 全新主题系列: 使用用户手动配置的多语言ID ${manualConfig.id}`);
        } else {
            // 全新系列但用户未配置
            result.multiLangConfig = null;
            result.source = 'none';
            console.log(`⚠️ 全新主题系列但缺少多语言配置`);
        }

        return result;
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
                    "RC现在的主题通道": "G3",
                    "作用": "装饰颜色1",
                    "颜色代码": "G1"
                },
                {
                    "RC现在的主题通道": "G1",
                    "作用": "装饰颜色2",
                    "颜色代码": "G2"
                },
                {
                    "RC现在的主题通道": "G2",
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

        // 使用新的文件选择状态更新函数，保持与其他文件选择的一致性
        const fileInfo = `文件名: ${data.fileName} | 大小: ${formatFileSize(data.fileSize || 0)} | 选择时间: ${getCurrentTimeString()}`;
        updateFileSelectionStatus('sourceFileStatus', 'success', '源数据文件选择成功', fileInfo);

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
     * 从RSC_Theme.xls中提取现有主题名称列表
     * @returns {Array} 现有主题名称数组
     */
    function extractExistingThemes() {
        console.log('=== 开始提取现有主题列表 ===');

        if (!rscThemeData || !rscThemeData.data || rscThemeData.data.length < 6) {
            console.warn('RSC_Theme数据未加载或数据不足（需要至少6行数据）');
            return [];
        }

        const data = rscThemeData.data;
        const headerRow = data[0];
        const notesColumnIndex = headerRow.findIndex(col => col === 'notes');

        if (notesColumnIndex === -1) {
            console.warn('在RSC_Theme.xls中未找到notes列');
            return [];
        }

        console.log(`notes列索引: ${notesColumnIndex}`);

        // 调试：输出前几行数据以验证结构
        console.log('RSC_Theme数据结构验证:');
        for (let i = 0; i < Math.min(8, data.length); i++) {
            const noteValue = data[i][notesColumnIndex];
            console.log(`第${i + 1}行 notes列值:`, noteValue);
        }

        // 从第6行开始提取主题名称（索引5），保持原始顺序
        const existingThemes = [];
        for (let i = 5; i < data.length; i++) {
            const themeName = data[i][notesColumnIndex];
            if (themeName && themeName.trim() !== '') {
                const trimmedName = themeName.trim();
                // 避免重复，但保持第一次出现的顺序
                if (!existingThemes.includes(trimmedName)) {
                    existingThemes.push(trimmedName);
                }
            }
        }

        console.log(`从第6行开始提取到 ${existingThemes.length} 个现有主题:`, existingThemes);
        console.log('=== 主题列表提取完成（保持原始顺序） ===');

        return existingThemes; // 保持表格中的原始排列顺序，不进行排序
    }

    /**
     * 更新主题选择器的选项列表
     */
    function updateThemeSelector() {
        const themeSelector = document.getElementById('themeSelector');
        if (!themeSelector) {
            console.warn('主题选择器元素未找到');
            return;
        }

        // 清空现有选项
        themeSelector.innerHTML = '<option value="">选择现有主题进行更新...</option>';

        // 获取现有主题列表
        const existingThemes = extractExistingThemes();

        // 添加主题选项
        existingThemes.forEach(themeName => {
            const option = document.createElement('option');
            option.value = themeName;
            option.textContent = themeName;
            themeSelector.appendChild(option);
        });

        console.log(`主题选择器已更新，包含 ${existingThemes.length} 个选项`);
    }

    /**
     * 处理主题选择器变化事件
     */
    function handleThemeSelection() {
        if (!themeSelector) return;

        const selectedTheme = themeSelector.value;

        if (selectedTheme) {
            // 清空文本输入框
            if (themeNameInput) {
                themeNameInput.value = '';
            }

            // 更新操作模式为更新模式
            updateOperationMode('update', selectedTheme);

            // 隐藏多语言配置面板（更新模式不需要）
            toggleMultiLangPanel(false);

            // 显示UGC配置面板并加载现有值
            toggleUGCConfigPanel(true);
            loadExistingUGCConfig(selectedTheme);

            // 显示Light配置面板并加载现有值
            toggleLightConfigPanel(true);
            loadExistingLightConfig(selectedTheme);

            // 显示ColorInfo配置面板并加载现有值
            toggleColorInfoConfigPanel(true);
            loadExistingColorInfoConfig(selectedTheme);

            // 启用处理按钮
            if (processThemeBtn) {
                processThemeBtn.disabled = false;
            }

            // 隐藏验证提示
            hideValidationMessage();
        } else {
            // 重置操作模式
            updateOperationMode('neutral');

            // 隐藏多语言配置面板
            toggleMultiLangPanel(false);

            // 隐藏UGC配置面板
            toggleUGCConfigPanel(false);

            // 禁用处理按钮
            if (processThemeBtn) {
                processThemeBtn.disabled = true;
            }
        }
    }

    /**
     * 处理主题输入框变化事件
     */
    function handleThemeInput() {
        if (!themeNameInput) return;

        const inputValue = themeNameInput.value.trim();

        if (inputValue) {
            // 清空选择器
            if (themeSelector) {
                themeSelector.value = '';
            }

            // 更新操作模式为创建模式
            updateOperationMode('create', inputValue);

            // 智能检测主题类型并决定是否显示多语言配置面板
            const smartConfig = getSmartMultiLanguageConfig(inputValue);
            toggleMultiLangPanel(smartConfig.shouldShowConfig);

            // 显示UGC配置面板（新建主题时总是显示）
            toggleUGCConfigPanel(true);

            // 显示Light配置面板（新建主题时总是显示）
            toggleLightConfigPanel(true);
            // 新建主题时使用最后一个主题的Light配置作为默认值
            resetLightConfigToDefaults();

            // 显示ColorInfo配置面板（新建主题时总是显示）
            toggleColorInfoConfigPanel(true);
            // 新建主题时使用最后一个主题的ColorInfo配置作为默认值
            resetColorInfoConfigToDefaults();

            // 更新多语言配置状态提示
            updateThemeTypeIndicator(smartConfig);

            // 验证输入
            validateThemeInput();
        } else {
            // 重置操作模式
            updateOperationMode('neutral');

            // 隐藏多语言配置面板
            toggleMultiLangPanel(false);

            // 隐藏UGC配置面板
            toggleUGCConfigPanel(false);

            // 隐藏Light配置面板
            toggleLightConfigPanel(false);

            // 隐藏ColorInfo配置面板
            toggleColorInfoConfigPanel(false);

            // 清除主题类型提示
            clearThemeTypeIndicator();

            // 禁用处理按钮
            if (processThemeBtn) {
                processThemeBtn.disabled = true;
            }

            // 隐藏验证提示
            hideValidationMessage();
        }
    }

    /**
     * 验证主题输入
     */
    function validateThemeInput() {
        if (!themeNameInput) return;

        const inputValue = themeNameInput.value.trim();

        if (!inputValue) {
            hideValidationMessage();
            if (processThemeBtn) {
                processThemeBtn.disabled = true;
            }
            return;
        }

        // 检查是否与现有主题重复
        const existingThemes = extractExistingThemes();
        const isDuplicate = existingThemes.some(theme =>
            theme.toLowerCase() === inputValue.toLowerCase()
        );

        if (isDuplicate) {
            showValidationMessage('error', `主题 "${inputValue}" 已存在，请选择更新现有主题或使用不同的名称`);
            if (processThemeBtn) {
                processThemeBtn.disabled = true;
            }
        } else {
            showValidationMessage('success', `将创建新主题 "${inputValue}"`);
            if (processThemeBtn) {
                processThemeBtn.disabled = false;
            }
        }
    }

    /**
     * 更新操作模式显示
     */
    function updateOperationMode(mode, themeName = '') {
        if (!modeBadge || !modeDescription) return;

        switch (mode) {
            case 'update':
                modeBadge.textContent = '🔄 更新模式';
                modeBadge.className = 'mode-badge update-mode';
                modeDescription.textContent = `将更新现有主题 "${themeName}" 的颜色数据`;
                break;
            case 'create':
                modeBadge.textContent = '✨ 创建模式';
                modeBadge.className = 'mode-badge create-mode';
                modeDescription.textContent = `将创建新主题 "${themeName}" 并添加到表格末尾`;
                break;
            default:
                modeBadge.textContent = '请选择操作模式';
                modeBadge.className = 'mode-badge neutral';
                modeDescription.textContent = '选择现有主题进行更新，或输入新主题名称进行创建';
                break;
        }
    }

    /**
     * 显示验证消息
     */
    function showValidationMessage(type, message) {
        if (!themeInputValidation) return;

        themeInputValidation.textContent = message;
        themeInputValidation.className = `input-validation ${type}`;
        themeInputValidation.style.display = 'block';
    }

    /**
     * 隐藏验证消息
     */
    function hideValidationMessage() {
        if (!themeInputValidation) return;

        themeInputValidation.style.display = 'none';
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

            // 更新主题选择器
            updateThemeSelector();

            App.Utils.showStatus('所有文件已准备就绪，请选择或输入主题名称', 'success');
        }
    }

    /**
     * 获取当前选择的主题名称和操作模式
     */
    function getCurrentThemeSelection() {
        let themeName = '';
        let operationMode = 'create';

        // 优先检查选择器
        if (themeSelector && themeSelector.value) {
            themeName = themeSelector.value.trim();
            operationMode = 'update';
        }
        // 其次检查输入框
        else if (themeNameInput && themeNameInput.value.trim()) {
            themeName = themeNameInput.value.trim();
            operationMode = 'create';
        }

        return { themeName, operationMode };
    }

    /**
     * 处理主题数据
     */
    async function processThemeData() {
        const { themeName, operationMode } = getCurrentThemeSelection();

        if (!themeName) {
            App.Utils.showStatus('请选择现有主题或输入新主题名称', 'error');
            return;
        }

        // 检查必要文件并提供详细的错误信息
        const missingFiles = [];
        if (!sourceData) missingFiles.push('源数据文件（包含完整配色表的Excel文件）');
        if (!rscThemeData) missingFiles.push('RSC_Theme.xls文件');
        if (!ugcThemeData) missingFiles.push('UGCTheme.xls文件');
        if (!mappingData) missingFiles.push('对比映射数据');

        if (missingFiles.length > 0) {
            const errorMessage = `缺少以下必要文件：\n• ${missingFiles.join('\n• ')}`;
            App.Utils.showStatus(errorMessage, 'error');
            console.warn('缺少必要文件:', missingFiles);
            return;
        }

        // 如果是创建新主题，使用智能检测进行多语言配置验证
        if (operationMode === 'create') {
            const smartConfig = getSmartMultiLanguageConfig(themeName);
            console.log('智能多语言配置检测结果:', smartConfig);

            if (smartConfig.isNewSeries && (!smartConfig.multiLangConfig || !smartConfig.multiLangConfig.isValid)) {
                // 全新主题系列但缺少多语言配置
                let errorMessage = '创建全新主题系列需要完整的多语言配置：\n\n';
                errorMessage += '• 缺少多语言ID\n';
                errorMessage += '\n请完成以下步骤：\n';
                errorMessage += '1. 点击"打开在线多语言表"按钮\n';
                errorMessage += '2. 在在线表中添加主题信息\n';
                errorMessage += '3. 回来输入分配的多语言ID';

                App.Utils.showStatus(errorMessage, 'error');

                // 高亮多语言ID输入框
                const multiLangIdInput = document.getElementById('multiLangId');
                if (multiLangIdInput) {
                    multiLangIdInput.focus();
                    multiLangIdInput.style.border = '2px solid #f44336';
                }

                return;
            } else if (!smartConfig.isNewSeries) {
                // 同系列主题，显示自动复用信息
                App.Utils.showStatus(`检测到同系列主题，将自动复用 "${smartConfig.similarity.matchedTheme}" 的多语言配置`, 'info', 3000);
                console.log('同系列主题验证通过，将自动复用多语言配置');
            } else {
                // 全新系列且配置完整
                console.log('全新主题系列多语言配置验证通过:', smartConfig.multiLangConfig);
            }
        }

        // 显示操作确认
        const confirmMessage = operationMode === 'update'
            ? `确认更新现有主题 "${themeName}" 吗？这将覆盖该主题的所有颜色数据。`
            : `确认创建新主题 "${themeName}" 吗？`;

        if (!confirm(confirmMessage)) {
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

        // 遍历映射数据 - 新逻辑：优先检查颜色代码
        mappingData.data.forEach((mapping, index) => {
            const rcChannel = mapping['RC现在的主题通道'];
            const colorCode = mapping['颜色代码'];

            console.log(`\n处理映射 ${index + 1}:`, {
                rcChannel,
                colorCode,
                作用: mapping['作用']
            });

            // 新逻辑：首先检查颜色代码是否存在
            if (!colorCode || colorCode === '' || colorCode === null || colorCode === undefined) {
                console.log(`跳过映射: 颜色代码为空 (${colorCode})`);
                return;
            }

            // 检查RC通道是否有效
            const isValidRCChannel = rcChannel &&
                                   rcChannel !== '' &&
                                   rcChannel !== '占不导入' &&
                                   rcChannel !== '暂不导入' &&
                                   rcChannel !== null &&
                                   rcChannel !== undefined;

            console.log(`颜色代码 ${colorCode} 对应的RC通道: ${rcChannel}, 有效性: ${isValidRCChannel}`);

            summary.total++;

            try {
                // 从源数据中查找对应的颜色值
                const colorValue = findColorValue(colorCode);
                console.log(`源数据中查找颜色代码 ${colorCode} 的结果: ${colorValue}`);

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

                // 新逻辑：根据RC通道有效性决定处理方式
                let columnIndex = -1;
                if (isValidRCChannel) {
                    // RC通道有效，查找对应列
                    columnIndex = headerRow.findIndex(col => col === rcChannel);

                    if (columnIndex === -1) {
                        const error = `未找到列: ${rcChannel}`;
                        console.error(error);
                        summary.errors.push(error);
                        return;
                    }

                    console.log(`找到目标列: ${rcChannel} (索引: ${columnIndex})`);
                } else {
                    // RC通道无效，记录但不更新数据
                    console.log(`RC通道无效 (${rcChannel})，颜色代码 ${colorCode} 使用默认处理`);
                    finalColorValue = 'FFFFFF';
                    isDefault = true;
                }

                // 根据RC通道有效性决定是否更新数据
                if (isValidRCChannel && columnIndex !== -1) {

                    // 确保数据更新到正确的位置
                    if (themeRow && columnIndex >= 0 && columnIndex < themeRow.length) {
                        themeRow[columnIndex] = finalColorValue;
                        console.log(`📝 数据更新: 行${rowIndex}, 列${columnIndex}(${rcChannel}) = ${finalColorValue}`);

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
                    } else {
                        console.error(`❌ 数据更新失败: 无效的行或列索引 - 行:${rowIndex}, 列:${columnIndex}`);
                        throw new Error(`无效的数据位置: 行${rowIndex}, 列${columnIndex}`);
                    }
                } else {
                    // RC通道无效，只记录但不更新实际数据
                    console.log(`🔄 跳过数据更新: RC通道无效 (${rcChannel}), 颜色代码: ${colorCode}`);

                    // 记录跳过的项目（用于统计）
                    updatedColors.push({
                        channel: rcChannel || '无效通道',
                        colorCode: colorCode,
                        value: finalColorValue,
                        isDefault: true,
                        skipped: true,
                        reason: 'RC通道无效'
                    });

                    summary.notFound++;
                }
            } catch (error) {
                const errorMsg = `处理${rcChannel}时出错: ${error.message}`;
                console.error(errorMsg, error);
                summary.errors.push(errorMsg);
            }
        });

        console.log('\n=== 颜色映射处理完成 ===');
        console.log('处理统计:', summary);
        console.log('有效映射数量:', updatedColors.filter(c => !c.skipped).length);
        console.log('跳过映射数量:', updatedColors.filter(c => c.skipped).length);
        console.log('成功更新数量:', summary.updated);
        console.log('使用默认值数量:', summary.notFound);
        console.log('错误数量:', summary.errors.length);

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
            console.log('更新现有主题，开始处理ColorInfo和Light sheet配置');
            return updateExistingThemeAdditionalSheets(themeName);
        }

        if (!rscThemeData || !rscThemeData.workbook) {
            console.error('RSC_Theme数据未加载');
            return { success: false, error: 'RSC_Theme数据未加载' };
        }

        try {
            const workbook = rscThemeData.workbook;
            const sheetNames = workbook.SheetNames;
            console.log('RSC_Theme包含的sheet:', sheetNames);

            // 严格限制：仅处理这3个目标工作表（主工作表、Light、ColorInfo）
            // 不影响RSC_Theme.xls文件中的其他工作表
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

                            // 更新rscAllSheetsData（关键修复：确保generateUpdatedWorkbook使用最新数据）
                            if (rscAllSheetsData) {
                                rscAllSheetsData[sheetName] = sheetData;
                                console.log(`✅ 已更新rscAllSheetsData["${sheetName}"]，新行数: ${sheetData.length}`);
                            }

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
     * 更新现有主题的Light和ColorInfo配置
     * @param {string} themeName - 主题名称
     * @returns {Object} 处理结果
     */
    function updateExistingThemeAdditionalSheets(themeName) {
        console.log('=== 开始更新现有主题的Light和ColorInfo配置 ===');
        console.log('主题名称:', themeName);

        if (!rscThemeData || !rscThemeData.workbook) {
            console.error('RSC_Theme数据未加载');
            return { success: false, error: 'RSC_Theme数据未加载' };
        }

        try {
            const workbook = rscThemeData.workbook;
            const sheetNames = workbook.SheetNames;
            console.log('RSC_Theme包含的sheet:', sheetNames);

            const targetSheets = ['ColorInfo', 'Light'];
            const updatedSheets = [];

            targetSheets.forEach(sheetName => {
                if (sheetNames.includes(sheetName)) {
                    console.log(`开始更新sheet: ${sheetName}`);

                    const worksheet = workbook.Sheets[sheetName];
                    const sheetData = XLSX.utils.sheet_to_json(worksheet, {
                        header: 1,
                        defval: '',
                        raw: false
                    });

                    if (sheetData.length > 0) {
                        const result = updateExistingRowInSheet(sheetData, themeName, sheetName);
                        if (result.success) {
                            // 更新工作表
                            const newWorksheet = XLSX.utils.aoa_to_sheet(sheetData);
                            workbook.Sheets[sheetName] = newWorksheet;

                            // 更新rscAllSheetsData
                            if (rscAllSheetsData) {
                                rscAllSheetsData[sheetName] = sheetData;
                            }

                            updatedSheets.push({
                                sheetName: sheetName,
                                rowIndex: result.rowIndex,
                                updated: true
                            });
                            console.log(`✅ ${sheetName} sheet更新成功`);
                        } else {
                            console.warn(`⚠️ ${sheetName} sheet更新失败:`, result.error);
                        }
                    } else {
                        console.warn(`⚠️ ${sheetName} sheet为空，跳过处理`);
                    }
                } else {
                    console.log(`Sheet "${sheetName}" 不存在，跳过处理`);
                }
            });

            console.log('现有主题Light和ColorInfo配置更新完成，更新的sheets:', updatedSheets);

            return {
                success: true,
                action: 'update_existing_rows',
                updatedSheets: updatedSheets,
                workbook: workbook
            };

        } catch (error) {
            console.error('更新现有主题Light和ColorInfo配置失败:', error);
            return {
                success: false,
                error: error.message
            };
        }
    }

    /**
     * 更新指定sheet中现有主题的行
     * @param {Array} sheetData - sheet数据数组
     * @param {string} themeName - 主题名称
     * @param {string} sheetName - sheet名称
     * @returns {Object} 处理结果
     */
    function updateExistingRowInSheet(sheetData, themeName, sheetName) {
        console.log(`=== 开始更新${sheetName}中的现有主题行 ===`);

        if (sheetData.length === 0) {
            return { success: false, error: 'Sheet数据为空' };
        }

        const headerRow = sheetData[0];
        console.log(`${sheetName} 表头:`, headerRow);

        // 查找notes列的索引
        const notesColumnIndex = headerRow.findIndex(col => col === 'notes');
        if (notesColumnIndex === -1) {
            console.warn(`${sheetName} 中未找到notes列`);
            return { success: false, error: 'notes列未找到' };
        }

        // 查找主题对应的行
        let themeRowIndex = -1;
        for (let i = 1; i < sheetData.length; i++) {
            if (sheetData[i][notesColumnIndex] === themeName) {
                themeRowIndex = i;
                break;
            }
        }

        if (themeRowIndex === -1) {
            console.log(`在${sheetName}中未找到主题"${themeName}"，跳过更新`);
            return { success: false, error: `主题"${themeName}"未找到` };
        }

        console.log(`在${sheetName}中找到主题"${themeName}"，行索引: ${themeRowIndex}`);

        // 获取现有行数据
        const existingRow = sheetData[themeRowIndex];
        console.log(`现有行数据:`, existingRow);

        // 根据sheet类型应用用户配置的数据
        if (sheetName === 'Light') {
            applyLightConfigToRow(headerRow, existingRow);
        } else if (sheetName === 'ColorInfo') {
            applyColorInfoConfigToRow(headerRow, existingRow);
        }

        console.log(`✅ ${sheetName}中主题"${themeName}"的配置已更新`);
        console.log(`更新后的行数据:`, existingRow);

        return {
            success: true,
            rowIndex: themeRowIndex,
            updatedRow: existingRow
        };
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

        // 根据sheet类型应用用户配置的数据
        if (sheetName === 'Light') {
            applyLightConfigToRow(headerRow, newRow);
        } else if (sheetName === 'ColorInfo') {
            applyColorInfoConfigToRow(headerRow, newRow);
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
     * 应用Light配置数据到新行
     * @param {Array} headerRow - 表头行
     * @param {Array} newRow - 新行数据
     */
    function applyLightConfigToRow(headerRow, newRow) {
        console.log('=== 开始应用Light配置数据到新行 ===');

        try {
            // 获取用户配置的Light数据
            const lightConfig = getLightConfigData();
            console.log('用户配置的Light数据:', lightConfig);

            // Light字段映射
            const lightFieldMapping = {
                'Max': 'Max',
                'Dark': 'Dark',
                'Min': 'Min',
                'SpecularLevel': 'SpecularLevel',
                'Gloss': 'Gloss',
                'SpecularColor': 'SpecularColor'
            };

            // 应用Light配置到新行
            Object.entries(lightFieldMapping).forEach(([columnName, configKey]) => {
                const columnIndex = headerRow.findIndex(col => col === columnName);
                if (columnIndex !== -1) {
                    const value = lightConfig[configKey];
                    newRow[columnIndex] = value.toString();
                    console.log(`Light配置: ${columnName} = ${value} (列索引: ${columnIndex})`);
                } else {
                    console.warn(`Light sheet中找不到列: ${columnName}`);
                }
            });

            console.log('✅ Light配置数据应用完成');
        } catch (error) {
            console.error('应用Light配置数据失败:', error);
        }
    }

    /**
     * 应用ColorInfo配置数据到新行
     * @param {Array} headerRow - 表头行
     * @param {Array} newRow - 新行数据
     */
    function applyColorInfoConfigToRow(headerRow, newRow) {
        console.log('=== 开始应用ColorInfo配置数据到新行 ===');

        try {
            // 获取用户配置的ColorInfo数据
            const colorInfoConfig = getColorInfoConfigData();
            console.log('用户配置的ColorInfo数据:', colorInfoConfig);

            // ColorInfo字段映射
            const colorInfoFieldMapping = {
                'PickupDiffR': 'PickupDiffR',
                'PickupDiffG': 'PickupDiffG',
                'PickupDiffB': 'PickupDiffB',
                'PickupReflR': 'PickupReflR',
                'PickupReflG': 'PickupReflG',
                'PickupReflB': 'PickupReflB',
                'BallSpecR': 'BallSpecR',
                'BallSpecG': 'BallSpecG',
                'BallSpecB': 'BallSpecB',
                'ForegroundFogR': 'ForegroundFogR',
                'ForegroundFogG': 'ForegroundFogG',
                'ForegroundFogB': 'ForegroundFogB',
                'FogStart': 'FogStart',
                'FogEnd': 'FogEnd'
            };

            // 应用ColorInfo配置到新行
            Object.entries(colorInfoFieldMapping).forEach(([columnName, configKey]) => {
                const columnIndex = headerRow.findIndex(col => col === columnName);
                if (columnIndex !== -1) {
                    const value = colorInfoConfig[configKey];
                    newRow[columnIndex] = value.toString();
                    console.log(`ColorInfo配置: ${columnName} = ${value} (列索引: ${columnIndex})`);
                } else {
                    console.warn(`ColorInfo sheet中找不到列: ${columnName}`);
                }
            });

            console.log('✅ ColorInfo配置数据应用完成');
        } catch (error) {
            console.error('应用ColorInfo配置数据失败:', error);
        }
    }

    /**
     * 从现有主题数据中查找多语言ID
     */
    function findMultiLangIdFromExistingTheme(baseThemeName, sheetData, levelNameColumnIndex) {
        if (!baseThemeName || !sheetData || levelNameColumnIndex === -1) {
            return null;
        }

        // 检查RSC数据是否可用
        if (!rscAllSheetsData || !rscAllSheetsData['Color']) {
            console.warn('RSC_Theme Color数据未加载，无法查找基础主题的多语言ID');
            return null;
        }

        // 第一步：在RSC_Theme的Color表中找到基础主题的行索引
        const rscColorData = rscAllSheetsData['Color'];
        const rscHeaderRow = rscColorData[0];
        const rscNotesColumnIndex = rscHeaderRow.findIndex(col => col === 'notes');

        if (rscNotesColumnIndex === -1) {
            console.warn('RSC_Theme Color表中找不到notes列，无法查找基础主题的多语言ID');
            return null;
        }

        // 查找基础主题在RSC中的行索引
        const rscThemeRowIndex = rscColorData.findIndex((row, index) =>
            index > 0 && row[rscNotesColumnIndex] === baseThemeName.trim()
        );

        if (rscThemeRowIndex === -1) {
            console.warn(`在RSC_Theme Color表中未找到基础主题 "${baseThemeName}"`);
            return null;
        }

        console.log(`在RSC_Theme Color表中找到基础主题 "${baseThemeName}"，行索引: ${rscThemeRowIndex}`);

        // 第二步：使用行索引在UGCTheme目标表中获取对应行的多语言ID
        if (rscThemeRowIndex < sheetData.length) {
            const targetRow = sheetData[rscThemeRowIndex];
            const multiLangId = targetRow[levelNameColumnIndex];
            if (multiLangId) {
                console.log(`找到基础主题 "${baseThemeName}" 的多语言ID: ${multiLangId} (行索引: ${rscThemeRowIndex})`);
                return multiLangId.toString();
            }
        }

        console.warn(`未找到基础主题 "${baseThemeName}" 的多语言ID (行索引: ${rscThemeRowIndex})`);
        return null;
    }

    /**
     * 从现有主题数据中查找Level_id
     */
    function findLevelIdFromExistingTheme(baseThemeName, sheetData, levelIdColumnIndex) {
        if (!baseThemeName || !sheetData || levelIdColumnIndex === -1) {
            return null;
        }

        // 检查RSC数据是否可用
        if (!rscAllSheetsData || !rscAllSheetsData['Color']) {
            console.warn('RSC_Theme Color数据未加载，无法查找基础主题的Level_id');
            return null;
        }

        // 第一步：在RSC_Theme的Color表中找到基础主题的行索引
        const rscColorData = rscAllSheetsData['Color'];
        const rscHeaderRow = rscColorData[0];
        const rscNotesColumnIndex = rscHeaderRow.findIndex(col => col === 'notes');

        if (rscNotesColumnIndex === -1) {
            console.warn('RSC_Theme Color表中找不到notes列，无法查找基础主题的Level_id');
            return null;
        }

        // 查找基础主题在RSC中的行索引
        const rscThemeRowIndex = rscColorData.findIndex((row, index) =>
            index > 0 && row[rscNotesColumnIndex] === baseThemeName.trim()
        );

        if (rscThemeRowIndex === -1) {
            console.warn(`在RSC_Theme Color表中未找到基础主题 "${baseThemeName}"`);
            return null;
        }

        console.log(`在RSC_Theme Color表中找到基础主题 "${baseThemeName}"，行索引: ${rscThemeRowIndex}`);

        // 第二步：使用行索引在UGCTheme目标表中获取对应行的Level_id
        if (rscThemeRowIndex < sheetData.length) {
            const targetRow = sheetData[rscThemeRowIndex];
            const levelId = targetRow[levelIdColumnIndex];
            if (levelId) {
                console.log(`找到基础主题 "${baseThemeName}" 的Level_id: ${levelId} (行索引: ${rscThemeRowIndex})`);
                return levelId.toString();
            }
        }

        console.warn(`未找到基础主题 "${baseThemeName}" 的Level_id (行索引: ${rscThemeRowIndex})`);
        return null;
    }

    /**
     * 从现有主题数据中查找Level_show_bg_ID
     */
    function findLevelShowBgIdFromExistingTheme(baseThemeName, sheetData, levelShowBgIdColumnIndex) {
        if (!baseThemeName || !sheetData || levelShowBgIdColumnIndex === -1) {
            return null;
        }

        // 检查RSC数据是否可用
        if (!rscAllSheetsData || !rscAllSheetsData['Color']) {
            console.warn('RSC_Theme Color数据未加载，无法查找基础主题的Level_show_bg_ID');
            return null;
        }

        // 第一步：在RSC_Theme的Color表中找到基础主题的行索引
        const rscColorData = rscAllSheetsData['Color'];
        const rscHeaderRow = rscColorData[0];
        const rscNotesColumnIndex = rscHeaderRow.findIndex(col => col === 'notes');

        if (rscNotesColumnIndex === -1) {
            console.warn('RSC_Theme Color表中找不到notes列，无法查找基础主题的Level_show_bg_ID');
            return null;
        }

        // 查找基础主题在RSC中的行索引
        const rscThemeRowIndex = rscColorData.findIndex((row, index) =>
            index > 0 && row[rscNotesColumnIndex] === baseThemeName.trim()
        );

        if (rscThemeRowIndex === -1) {
            console.warn(`在RSC_Theme Color表中未找到基础主题 "${baseThemeName}"`);
            return null;
        }

        console.log(`在RSC_Theme Color表中找到基础主题 "${baseThemeName}"，行索引: ${rscThemeRowIndex}`);

        // 第二步：使用行索引在UGCTheme目标表中获取对应行的Level_show_bg_ID
        if (rscThemeRowIndex < sheetData.length) {
            const targetRow = sheetData[rscThemeRowIndex];
            const levelShowBgId = targetRow[levelShowBgIdColumnIndex];
            if (levelShowBgId !== undefined && levelShowBgId !== null) {
                console.log(`找到基础主题 "${baseThemeName}" 的Level_show_bg_ID: ${levelShowBgId} (行索引: ${rscThemeRowIndex})`);
                return levelShowBgId.toString();
            }
        }

        console.warn(`未找到基础主题 "${baseThemeName}" 的Level_show_bg_ID (行索引: ${rscThemeRowIndex})`);
        return null;
    }

    /**
     * 查找同系列主题的最后一个主题ID
     */
    function findLastSimilarThemeId(baseThemeName, sheetData, idColumnIndex) {
        if (!baseThemeName || !sheetData || idColumnIndex === -1) {
            return null;
        }

        // 检查RSC数据是否可用
        if (!rscAllSheetsData || !rscAllSheetsData['Color']) {
            console.warn('RSC_Theme Color数据未加载，无法查找同系列主题ID');
            return null;
        }

        // 提取基础主题名称（去除数字）
        const baseName = extractThemeBaseName(baseThemeName);
        if (!baseName) {
            console.warn(`无法提取基础主题名称: ${baseThemeName}`);
            return null;
        }

        console.log(`查找同系列主题，基础名称: "${baseName}"`);

        // 在RSC_Theme Color表中查找所有同系列主题
        const rscColorData = rscAllSheetsData['Color'];
        const rscHeaderRow = rscColorData[0];
        const rscNotesColumnIndex = rscHeaderRow.findIndex(col => col === 'notes');

        if (rscNotesColumnIndex === -1) {
            console.warn('RSC_Theme Color表中找不到notes列');
            return null;
        }

        const similarThemes = [];
        for (let i = 1; i < rscColorData.length; i++) {
            const row = rscColorData[i];
            const themeName = row[rscNotesColumnIndex];
            if (themeName && extractThemeBaseName(themeName) === baseName) {
                // 获取对应的UGC主题ID
                if (i < sheetData.length) {
                    const ugcRow = sheetData[i];
                    const themeId = parseInt(ugcRow[idColumnIndex]) || 0;
                    if (themeId > 0) {
                        similarThemes.push({
                            name: themeName,
                            id: themeId,
                            rscIndex: i,
                            ugcIndex: i
                        });
                    }
                }
            }
        }

        if (similarThemes.length === 0) {
            console.log(`未找到同系列主题: ${baseName}`);
            return null;
        }

        // 按ID排序，找到最大的ID（最后一个主题）
        similarThemes.sort((a, b) => a.id - b.id);
        const lastTheme = similarThemes[similarThemes.length - 1];

        console.log(`找到 ${similarThemes.length} 个同系列主题:`, similarThemes.map(t => `${t.name}(ID:${t.id})`));
        console.log(`最后一个同系列主题: ${lastTheme.name} (ID: ${lastTheme.id})`);

        return lastTheme.id;
    }

    /**
     * 执行同系列主题的排序插入操作
     */
    function performSortedInsertion(data, headerRow, newThemeId, targetLevelShowId, themeName) {
        console.log('=== 开始执行排序插入操作 ===');
        console.log(`新主题ID: ${newThemeId}, 目标Level_show_id: ${targetLevelShowId}, 主题名称: ${themeName}`);

        // 查找Level_show_id列和M列的索引
        const levelShowIdColumnIndex = headerRow.findIndex(col => col === 'Level_show_id');
        const mColumnIndex = 12; // M列固定为索引12

        if (levelShowIdColumnIndex === -1) {
            console.error('未找到Level_show_id列，无法执行排序插入');
            return { success: false, error: '未找到Level_show_id列' };
        }

        console.log(`Level_show_id列索引: ${levelShowIdColumnIndex}, M列索引: ${mColumnIndex}`);

        // 查找目标Level_show_id值对应的行（从第6行开始查找）
        let targetRowIndex = -1;
        for (let i = 5; i < data.length; i++) { // 从第6行开始（索引5，跳过前5行）
            const row = data[i];
            const levelShowIdValue = parseInt(row[levelShowIdColumnIndex]) || 0;
            if (levelShowIdValue === targetLevelShowId) {
                targetRowIndex = i;
                break;
            }
        }

        console.log(`在Custom_Ground_Color中从第6行开始查找Level_show_id=${targetLevelShowId}`);

        if (targetRowIndex === -1) {
            console.warn(`未找到Level_show_id值为 ${targetLevelShowId} 的行，排序插入操作取消`);
            return { success: false, error: `未找到Level_show_id值为 ${targetLevelShowId} 的行` };
        }

        console.log(`找到目标插入位置: 行索引 ${targetRowIndex} (Level_show_id=${targetLevelShowId})`);

        // 执行数据下移操作：将目标行以下的所有行的Level_show_id和M列数据向下移动
        const movedRows = [];

        // 从最后一行开始向前处理，避免数据覆盖
        for (let i = data.length - 1; i > targetRowIndex; i--) {
            const currentRow = data[i];
            const prevRow = data[i - 1];

            const oldLevelShowId = currentRow[levelShowIdColumnIndex];
            const oldMValue = currentRow[mColumnIndex];

            // 将上一行的Level_show_id和M列数据复制到当前行
            currentRow[levelShowIdColumnIndex] = prevRow[levelShowIdColumnIndex];
            currentRow[mColumnIndex] = prevRow[mColumnIndex];

            movedRows.push({
                rowIndex: i,
                oldLevelShowId: oldLevelShowId,
                oldMValue: oldMValue,
                newLevelShowId: currentRow[levelShowIdColumnIndex],
                newMValue: currentRow[mColumnIndex]
            });
        }

        console.log(`数据下移完成，影响 ${movedRows.length} 行:`, movedRows);

        // 在目标行的下一行插入新主题数据
        const insertRowIndex = targetRowIndex + 1;
        if (insertRowIndex < data.length) {
            const insertRow = data[insertRowIndex];
            insertRow[levelShowIdColumnIndex] = (newThemeId - 1).toString();
            insertRow[mColumnIndex] = themeName;

            console.log(`新主题数据已插入到行索引 ${insertRowIndex}:`);
            console.log(`  Level_show_id: ${insertRow[levelShowIdColumnIndex]}`);
            console.log(`  M列: ${insertRow[mColumnIndex]}`);
        } else {
            console.error(`插入位置 ${insertRowIndex} 超出数据范围`);
            return { success: false, error: '插入位置超出数据范围' };
        }

        console.log('=== 排序插入操作完成 ===');
        return {
            success: true,
            targetRowIndex: targetRowIndex,
            insertRowIndex: insertRowIndex,
            movedRowsCount: movedRows.length,
            newLevelShowId: newThemeId - 1
        };
    }

    /**
     * 更新文件选择状态显示
     * @param {string} statusId - 状态指示器的ID
     * @param {string} type - 状态类型：'success', 'error', 'loading'
     * @param {string|object} messageOrOptions - 状态消息或选项对象
     * @param {string} info - 详细信息（可选，用于向后兼容）
     */
    function updateFileSelectionStatus(statusId, type, messageOrOptions, info = '') {
        const statusElement = document.getElementById(statusId);
        if (!statusElement) {
            // 静默处理，不显示警告（因为某些状态指示器可能已被移除）
            return;
        }

        // 处理参数格式：支持新格式（对象）和旧格式（字符串）
        let message, fileName, fileSize, errorMessage;

        if (typeof messageOrOptions === 'object' && messageOrOptions !== null) {
            // 新格式：options对象
            fileName = messageOrOptions.fileName;
            fileSize = messageOrOptions.fileSize;
            errorMessage = messageOrOptions.errorMessage;

            // 根据状态类型生成消息
            const currentTime = getCurrentTimeString();
            switch (type) {
                case 'success':
                    message = `✅ ${fileName} (${formatFileSize(fileSize)}) - 选择于 ${currentTime}`;
                    break;
                case 'error':
                    message = `❌ ${fileName} - ${errorMessage}`;
                    break;
                case 'loading':
                    message = `⏳ 正在处理 ${fileName} (${formatFileSize(fileSize)})...`;
                    break;
                default:
                    message = '未选择文件';
            }
        } else {
            // 旧格式：字符串消息
            message = messageOrOptions || '';
        }

        // 显示状态指示器
        statusElement.style.display = 'block';

        // 清除之前的状态类
        statusElement.classList.remove('success', 'error', 'loading');

        // 添加新的状态类
        statusElement.classList.add(type);

        // 设置状态图标和消息
        const statusIcon = statusElement.querySelector('.status-icon');
        const statusMessage = statusElement.querySelector('.status-message');
        const statusInfo = statusElement.querySelector('.status-info');

        if (statusIcon) {
            switch (type) {
                case 'success':
                    statusIcon.textContent = '✅';
                    break;
                case 'error':
                    statusIcon.textContent = '❌';
                    break;
                case 'loading':
                    statusIcon.textContent = '⏳';
                    break;
                default:
                    statusIcon.textContent = 'ℹ️';
            }
        }

        if (statusMessage) {
            statusMessage.textContent = message;
        }

        if (statusInfo) {
            statusInfo.textContent = info;
        }

        console.log(`文件选择状态已更新: ${statusId} - ${type} - ${message}`);
    }

    /**
     * 隐藏文件选择状态显示
     * @param {string} statusId - 状态指示器的ID
     */
    function hideFileSelectionStatus(statusId) {
        const statusElement = document.getElementById(statusId);
        if (statusElement) {
            statusElement.style.display = 'none';
            statusElement.classList.remove('success', 'error', 'loading');
        }
    }

    /**
     * 格式化文件大小
     * @param {number} bytes - 文件大小（字节）
     * @returns {string} 格式化后的文件大小
     */
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    /**
     * 获取当前时间字符串
     * @returns {string} 格式化的时间字符串
     */
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

    /**
     * 显示最终操作指引弹框
     */
    function showFinalGuideModal() {
        const modal = document.getElementById('finalGuideModal');
        if (!modal) {
            console.error('最终操作指引弹框元素未找到');
            return;
        }

        // 使用固定路径显示
        const toolsPath = 'rs_ugc2\\RS_UGC_Unity_2021.3.31\\Tools';

        // 更新弹框中的路径信息
        const pathElement = document.getElementById('toolsPathDisplay');
        if (pathElement) {
            pathElement.textContent = toolsPath;
        }

        // 显示弹框
        modal.style.display = 'flex';

        // 移除自动关闭功能，隐藏倒计时元素
        const countdownElement = document.getElementById('autoCloseCountdown');
        if (countdownElement) {
            countdownElement.style.display = 'none';
        }

        // 绑定关闭事件（仅在用户点击按钮时关闭）
        const closeBtn = document.getElementById('closeFinalGuideModal');
        const okBtn = document.getElementById('finalGuideOkBtn');

        const closeHandler = () => {
            hideFinalGuideModal();
        };

        if (closeBtn) {
            closeBtn.onclick = closeHandler;
        }

        if (okBtn) {
            okBtn.onclick = closeHandler;
        }

        // 点击背景关闭
        modal.onclick = (e) => {
            if (e.target === modal) {
                closeHandler();
            }
        };

        console.log('最终操作指引弹框已显示，Tools路径:', toolsPath);
    }

    /**
     * 隐藏最终操作指引弹框
     */
    function hideFinalGuideModal() {
        const modal = document.getElementById('finalGuideModal');
        if (modal) {
            modal.style.display = 'none';

            // 清除事件监听器
            const closeBtn = document.getElementById('closeFinalGuideModal');
            const okBtn = document.getElementById('finalGuideOkBtn');

            if (closeBtn) closeBtn.onclick = null;
            if (okBtn) okBtn.onclick = null;
            modal.onclick = null;
        }

        console.log('最终操作指引弹框已隐藏');
    }

    /**
     * 读取Levels.xls文件并获取可用的levelid列表
     */
    async function loadLevelsData() {
        try {
            // 尝试从Unity项目文件中找到Levels.xls
            if (!unityProjectFiles || !unityProjectFiles.levelsFile) {
                console.warn('未找到Levels.xls文件，无法获取levelid数据');
                return null;
            }

            const levelsFileHandle = unityProjectFiles.levelsFile;
            console.log('开始读取Levels.xls文件...');

            // 获取文件内容
            const levelsFile = await levelsFileHandle.getFile();
            const arrayBuffer = await levelsFile.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });

            // 查找LevelInfo工作表
            const levelInfoSheetName = workbook.SheetNames.find(name =>
                name === 'LevelInfo' ||
                name.toLowerCase().includes('levelinfo') ||
                name.toLowerCase().includes('level_info')
            );

            if (!levelInfoSheetName) {
                console.warn('Levels.xls中未找到LevelInfo工作表');
                return null;
            }

            const worksheet = workbook.Sheets[levelInfoSheetName];
            const data = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                defval: '',
                raw: false
            });

            if (data.length < 6) {
                console.warn('LevelInfo工作表数据不足，至少需要6行数据');
                return null;
            }

            // 查找levelid列
            const headerRow = data[0];
            const levelIdColumnIndex = headerRow.findIndex(col =>
                col === 'levelid' ||
                col === 'LevelId' ||
                col === 'level_id' ||
                col.toLowerCase().includes('levelid')
            );

            if (levelIdColumnIndex === -1) {
                console.warn('LevelInfo工作表中未找到levelid列');
                return null;
            }

            // 提取从第6行开始的levelid数据
            const levelIds = [];
            for (let i = 5; i < data.length; i++) { // 从第6行开始（索引5）
                const row = data[i];
                const levelId = row[levelIdColumnIndex];
                if (levelId && levelId.toString().trim() !== '') {
                    const parsedId = parseInt(levelId);
                    if (!isNaN(parsedId)) {
                        levelIds.push(parsedId);
                    }
                }
            }

            console.log(`从Levels.xls加载了 ${levelIds.length} 个有效的levelid:`, levelIds);
            return {
                levelIds: levelIds,
                sheetName: levelInfoSheetName,
                columnName: headerRow[levelIdColumnIndex]
            };

        } catch (error) {
            console.error('读取Levels.xls文件失败:', error);
            return null;
        }
    }

    /**
     * 为全新主题系列随机选择Level_id
     */
    function selectRandomLevelId(currentLevelId, levelsData) {
        if (!levelsData || !levelsData.levelIds || levelsData.levelIds.length === 0) {
            console.warn('没有可用的levelid数据，无法随机选择');
            return null;
        }

        const currentId = parseInt(currentLevelId);

        // 过滤掉与当前Level_id相同的值
        const availableLevelIds = levelsData.levelIds.filter(id => id !== currentId);

        if (availableLevelIds.length === 0) {
            console.warn('所有levelid都与当前值相同，使用原始列表的第一个值');
            return levelsData.levelIds[0];
        }

        // 随机选择一个
        const randomIndex = Math.floor(Math.random() * availableLevelIds.length);
        const selectedLevelId = availableLevelIds[randomIndex];

        console.log(`从 ${availableLevelIds.length} 个可用levelid中随机选择: ${selectedLevelId} (排除了当前值: ${currentId})`);
        return selectedLevelId;
    }

    /**
     * 应用UGC字段设置到新行
     * @param {string} sheetName - Sheet名称
     * @param {Array} headerRow - 表头行
     * @param {Array} newRow - 新行数据
     */
    function applyUGCFieldSettings(sheetName, headerRow, newRow) {
        const ugcConfig = getUGCConfigData();
        console.log(`应用UGC字段设置到Sheet ${sheetName}:`, ugcConfig);

        // 定义每个sheet对应的字段映射
        const sheetFieldMapping = {
            'Custom_Ground_Color': {
                '_PatternUpIndex': ugcConfig.groundPatternIndex,
                '_FrameIndex': ugcConfig.groundFrameIndex
            },
            'Custom_Fragile_Color': {
                '_PatternUpIndex': ugcConfig.fragilePatternIndex,
                '_FrameIndex': ugcConfig.fragileFrameIndex,
                '_GlassAlpha': ugcConfig.fragileGlassAlpha,
                '_PatternAlpha': ugcConfig.fragilePatternAlpha
            },
            'Custom_Fragile_Active_Color': {
                '_PatternUpIndex': ugcConfig.fragileActivePatternIndex,
                '_FrameIndex': ugcConfig.fragileActiveFrameIndex,
                '_GlassAlpha': ugcConfig.fragileActiveGlassAlpha,
                '_PatternAlpha': ugcConfig.fragileActivePatternAlpha
            },
            'Custom_Jump_Color': {
                '_PatternUpIndex': ugcConfig.jumpPatternIndex,
                '_FrameIndex': ugcConfig.jumpFrameIndex
            },
            'Custom_Jump_Active_Color': {
                '_PatternUpIndex': ugcConfig.jumpActivePatternIndex,
                '_FrameIndex': ugcConfig.jumpActiveFrameIndex
            }
        };

        const fieldMapping = sheetFieldMapping[sheetName];
        if (!fieldMapping) {
            console.log(`Sheet ${sheetName} 不需要UGC字段设置`);
            return;
        }

        // 应用字段设置
        Object.entries(fieldMapping).forEach(([columnName, value]) => {
            const columnIndex = headerRow.findIndex(col => col === columnName);
            if (columnIndex !== -1) {
                newRow[columnIndex] = value.toString();
                console.log(`Sheet ${sheetName} 设置 ${columnName} = ${value} (列索引: ${columnIndex})`);
            } else {
                console.warn(`Sheet ${sheetName} 中找不到列 ${columnName}`);
            }
        });
    }

    /**
     * 更新现有主题的UGCTheme配置
     * @param {string} themeName - 主题名称
     * @returns {Object} 处理结果
     */
    async function updateExistingUGCTheme(themeName) {
        console.log('=== 开始更新现有主题的UGCTheme配置 ===');
        console.log('主题名称:', themeName);

        if (!ugcThemeData || !ugcThemeData.workbook) {
            console.error('UGCTheme数据未加载');
            return { success: false, error: 'UGCTheme数据未加载' };
        }

        if (!rscAllSheetsData) {
            console.error('RSC数据未加载');
            return { success: false, error: 'RSC数据未加载' };
        }

        try {
            // 第一步：在RSC_Theme的Color表中找到主题对应的行号
            const rscColorData = rscAllSheetsData['Color'];
            if (!rscColorData || rscColorData.length === 0) {
                console.error('RSC_Theme的Color表未找到或为空');
                return { success: false, error: 'RSC_Theme的Color表未找到' };
            }

            const rscHeaderRow = rscColorData[0];
            const rscNotesColumnIndex = rscHeaderRow.findIndex(col => col === 'notes');

            if (rscNotesColumnIndex === -1) {
                console.error('RSC_Theme的Color表没有notes列');
                return { success: false, error: 'RSC_Theme的Color表没有notes列' };
            }

            // 查找主题在RSC中的行号
            const rscThemeRowIndex = rscColorData.findIndex((row, index) =>
                index > 0 && row[rscNotesColumnIndex] === themeName
            );

            if (rscThemeRowIndex === -1) {
                console.error(`在RSC_Theme的Color表中未找到主题 "${themeName}"`);
                return { success: false, error: `主题 "${themeName}" 在RSC中不存在` };
            }

            console.log(`在RSC_Theme的Color表中找到主题 "${themeName}"，行索引: ${rscThemeRowIndex}`);
            const targetRowNumber = rscThemeRowIndex;

            // 第二步：获取用户配置的UGC数据
            const ugcConfig = getUGCConfigData();
            console.log('用户配置的UGC数据:', ugcConfig);

            const workbook = ugcThemeData.workbook;
            const processedSheets = [];

            // 定义需要更新的sheet和字段映射
            const sheetFieldMapping = {
                'Custom_Ground_Color': {
                    '_PatternUpIndex': ugcConfig.groundPatternIndex,
                    '_FrameIndex': ugcConfig.groundFrameIndex
                },
                'Custom_Fragile_Color': {
                    '_PatternUpIndex': ugcConfig.fragilePatternIndex,
                    '_FrameIndex': ugcConfig.fragileFrameIndex,
                    '_GlassAlpha': ugcConfig.fragileGlassAlpha,
                    '_PatternAlpha': ugcConfig.fragilePatternAlpha
                },
                'Custom_Fragile_Active_Color': {
                    '_PatternUpIndex': ugcConfig.fragileActivePatternIndex,
                    '_FrameIndex': ugcConfig.fragileActiveFrameIndex,
                    '_GlassAlpha': ugcConfig.fragileActiveGlassAlpha,
                    '_PatternAlpha': ugcConfig.fragileActivePatternAlpha
                },
                'Custom_Jump_Color': {
                    '_PatternUpIndex': ugcConfig.jumpPatternIndex,
                    '_FrameIndex': ugcConfig.jumpFrameIndex
                },
                'Custom_Jump_Active_Color': {
                    '_PatternUpIndex': ugcConfig.jumpActivePatternIndex,
                    '_FrameIndex': ugcConfig.jumpActiveFrameIndex
                }
            };

            // 第三步：更新每个相关的sheet
            Object.entries(sheetFieldMapping).forEach(([sheetName, fieldMapping]) => {
                console.log(`\n--- 更新Sheet: ${sheetName} ---`);

                const worksheet = workbook.Sheets[sheetName];
                if (!worksheet) {
                    console.log(`Sheet ${sheetName} 不存在，跳过`);
                    return;
                }

                const data = XLSX.utils.sheet_to_json(worksheet, {
                    header: 1,
                    defval: '',
                    raw: false
                });

                if (data.length === 0 || targetRowNumber >= data.length) {
                    console.log(`Sheet ${sheetName} 数据不足或目标行不存在，跳过`);
                    return;
                }

                const headerRow = data[0];
                const targetRow = data[targetRowNumber];

                console.log(`Sheet ${sheetName} 更新行 ${targetRowNumber}:`, targetRow);

                // 更新字段值
                Object.entries(fieldMapping).forEach(([columnName, value]) => {
                    const columnIndex = headerRow.findIndex(col => col === columnName);
                    if (columnIndex !== -1) {
                        const oldValue = targetRow[columnIndex];
                        targetRow[columnIndex] = value.toString();
                        console.log(`Sheet ${sheetName} 更新 ${columnName}: ${oldValue} -> ${value}`);
                    } else {
                        console.warn(`Sheet ${sheetName} 中找不到列 ${columnName}`);
                    }
                });

                // 更新worksheet
                const newWorksheet = XLSX.utils.aoa_to_sheet(data);
                workbook.Sheets[sheetName] = newWorksheet;

                processedSheets.push({
                    sheetName: sheetName,
                    updatedRowIndex: targetRowNumber,
                    updatedFields: Object.keys(fieldMapping)
                });
            });

            console.log('UGCTheme更新完成，处理的sheets:', processedSheets);

            // 同步UGC内存数据状态
            console.log('=== 开始同步UGC内存数据状态 ===');
            syncUGCMemoryDataState(workbook);

            return {
                success: true,
                action: 'update_existing',
                processedSheets: processedSheets,
                workbook: workbook
            };

        } catch (error) {
            console.error('更新现有UGCTheme失败:', error);
            return {
                success: false,
                error: error.message
            };
        }
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
            console.log('更新现有主题，开始更新UGCTheme配置');
            return await updateExistingUGCTheme(themeName);
        }

        if (!ugcThemeData || !ugcThemeData.workbook) {
            console.error('UGCTheme数据未加载');
            return { success: false, error: 'UGCTheme数据未加载' };
        }

        try {
            const workbook = ugcThemeData.workbook;
            const sheetNames = workbook.SheetNames;
            console.log('UGCTheme包含的sheet:', sheetNames);

            // 获取智能多语言配置
            const smartConfig = getSmartMultiLanguageConfig(themeName);
            console.log('UGCTheme处理 - 智能配置:', smartConfig);

            // 加载Levels数据（用于全新主题系列的Level_id随机选择）
            let levelsData = null;

            if (!smartConfig.similarity.isSimilar) {
                // 全新主题系列，需要加载Levels数据
                console.log('全新主题系列，开始加载Levels数据...');
                levelsData = await loadLevelsData();
                if (levelsData) {
                    console.log('Levels数据加载成功，准备随机选择Level_id');
                } else {
                    console.warn('Levels数据加载失败，将使用默认Level_id处理');
                }
            } else {
                console.log(`同系列主题，将复用基础主题 "${smartConfig.similarity.matchedTheme}" 的Level_id`);
            }

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

                // 获取智能多语言配置
                const smartConfig = getSmartMultiLanguageConfig(themeName);
                console.log(`Sheet ${sheetName} 智能多语言配置:`, smartConfig);

                // 声明变量用于记录处理结果
                let levelIdSource = 'unknown';

                // 处理多语言ID填充
                const levelNameColumnIndex = headerRow.findIndex(col =>
                    col === 'LevelName' ||
                    col === 'levelname' ||
                    col === 'LevelName_ID' ||
                    col === 'levelname_id' ||
                    col.toLowerCase().includes('levelname')
                );

                if (levelNameColumnIndex !== -1) {
                    const columnName = headerRow[levelNameColumnIndex];
                    let finalMultiLangId = null;
                    let source = 'unknown';

                    if (smartConfig.similarity.isSimilar) {
                        // 同系列主题，查找基础主题的多语言ID
                        const baseThemeMultiLangId = findMultiLangIdFromExistingTheme(smartConfig.similarity.matchedTheme, data, levelNameColumnIndex);
                        if (baseThemeMultiLangId) {
                            finalMultiLangId = baseThemeMultiLangId;
                            source = `auto_from_${smartConfig.similarity.matchedTheme}`;
                            console.log(`Sheet ${sheetName} 同系列主题，复用基础主题 "${smartConfig.similarity.matchedTheme}" 的多语言ID: ${finalMultiLangId}`);
                        } else {
                            // 找不到基础主题的多语言ID，使用上一行数据
                            finalMultiLangId = lastRow[levelNameColumnIndex] || '';
                            source = 'previous_row_fallback';
                            console.log(`Sheet ${sheetName} 无法找到基础主题的多语言ID，使用上一行数据: ${finalMultiLangId}`);
                        }
                    } else if (smartConfig.multiLangConfig && smartConfig.multiLangConfig.isValid && smartConfig.multiLangConfig.id) {
                        // 全新系列，使用用户输入的多语言ID
                        finalMultiLangId = smartConfig.multiLangConfig.id.toString();
                        source = 'user_input';
                        console.log(`Sheet ${sheetName} 全新主题系列，使用用户输入的多语言ID: ${finalMultiLangId}`);
                    } else {
                        // 使用上一行的数据作为默认值
                        finalMultiLangId = lastRow[levelNameColumnIndex] || '';
                        source = 'previous_row_default';
                        console.log(`Sheet ${sheetName} 多语言配置无效，使用上一行数据: ${finalMultiLangId}`);
                    }

                    newRow[levelNameColumnIndex] = finalMultiLangId;
                    console.log(`Sheet ${sheetName} 最终设置多语言ID: ${columnName} = ${finalMultiLangId} (来源: ${source})`);
                } else {
                    console.warn(`Sheet ${sheetName} 中找不到LevelName相关列`);
                }

                // 处理Level_show_bg_ID列的智能设置（同系列主题复用）
                let targetColumnIndex = headerRow.findIndex(col => col === 'Level_show_bg_ID');
                if (targetColumnIndex !== -1) {
                    let finalLevelShowBgId = null;
                    let levelShowBgIdSource = 'unknown';

                    if (smartConfig.similarity.isSimilar) {
                        // 同系列主题，复用基础主题的Level_show_bg_ID
                        const baseLevelShowBgId = findLevelShowBgIdFromExistingTheme(smartConfig.similarity.matchedTheme, data, targetColumnIndex);
                        if (baseLevelShowBgId) {
                            finalLevelShowBgId = baseLevelShowBgId;
                            levelShowBgIdSource = `auto_from_${smartConfig.similarity.matchedTheme}`;
                            console.log(`Sheet ${sheetName} 同系列主题，复用基础主题 "${smartConfig.similarity.matchedTheme}" 的Level_show_bg_ID: ${finalLevelShowBgId}`);
                        } else {
                            // 找不到基础主题的Level_show_bg_ID，使用默认值
                            finalLevelShowBgId = "-1";
                            levelShowBgIdSource = 'default_fallback';
                            console.log(`Sheet ${sheetName} 无法找到基础主题的Level_show_bg_ID，使用默认值: ${finalLevelShowBgId}`);
                        }
                    } else {
                        // 全新主题系列，使用默认值
                        finalLevelShowBgId = "-1";
                        levelShowBgIdSource = 'default_new_series';
                        console.log(`Sheet ${sheetName} 全新主题系列，使用默认Level_show_bg_ID: ${finalLevelShowBgId}`);
                    }

                    newRow[targetColumnIndex] = finalLevelShowBgId;
                    console.log(`Sheet ${sheetName} 最终设置Level_show_bg_ID: ${finalLevelShowBgId} (来源: ${levelShowBgIdSource})`);
                }else{
                    console.warn(`在${sheetName}中找不到Level_show_bg_ID列`);
                }

                // 处理Level_show_id列的智能设置（根据主题类型决定填值策略）
                targetColumnIndex = headerRow.findIndex(col => col === 'Level_show_id');
                if (targetColumnIndex !== -1) {
                    let finalLevelShowId = null;
                    let levelShowIdSource = 'unknown';

                    if (smartConfig.similarity.isSimilar) {
                        // 现有主题系列新增行，使用"上一行数据值+1"的逻辑
                        const currentValue = parseInt(lastRow[targetColumnIndex]) || 0;
                        finalLevelShowId = (currentValue + 1).toString();
                        levelShowIdSource = 'existing_series_increment';
                        console.log(`Sheet ${sheetName} 现有主题系列，Level_show_id递增: ${currentValue} + 1 = ${finalLevelShowId}`);
                    } else {
                        // 全新主题行，Level_show_id设置为"新增的主题ID - 1"
                        finalLevelShowId = (newId - 1).toString();
                        levelShowIdSource = 'new_theme_id_minus_one';
                        console.log(`Sheet ${sheetName} 全新主题系列，Level_show_id设置为新ID减1: ${newId} - 1 = ${finalLevelShowId}`);
                    }

                    newRow[targetColumnIndex] = finalLevelShowId;
                    console.log(`Sheet ${sheetName} 最终设置Level_show_id: ${finalLevelShowId} (来源: ${levelShowIdSource})`);
                } else {
                    console.warn(`在${sheetName}中找不到Level_show_id列`);
                }

                // 处理Level_id列的智能设置（同系列主题复用）
                const levelIdColumnIndex = headerRow.findIndex(col =>
                    col === 'Level_id' ||
                    col === 'level_id' ||
                    col === 'LevelId' ||
                    col.toLowerCase().includes('level_id')
                );

                if (levelIdColumnIndex !== -1) {
                    const levelIdColumnName = headerRow[levelIdColumnIndex];
                    let finalLevelId = null;

                    if (smartConfig.similarity.isSimilar) {
                        // 同系列主题，复用基础主题的Level_id
                        const baseLevelId = findLevelIdFromExistingTheme(smartConfig.similarity.matchedTheme, data, levelIdColumnIndex);
                        if (baseLevelId) {
                            finalLevelId = baseLevelId;
                            levelIdSource = `auto_from_${smartConfig.similarity.matchedTheme}`;
                            console.log(`Sheet ${sheetName} 同系列主题，复用基础主题 "${smartConfig.similarity.matchedTheme}" 的Level_id: ${finalLevelId}`);
                        } else {
                            // 找不到基础主题的Level_id，使用上一行数据
                            finalLevelId = lastRow[levelIdColumnIndex] || '1';
                            levelIdSource = 'previous_row_fallback';
                            console.log(`Sheet ${sheetName} 无法找到基础主题的Level_id，使用上一行数据: ${finalLevelId}`);
                        }
                    } else {
                        // 全新主题系列，从Levels文件随机选择Level_id
                        if (levelsData) {
                            const currentLevelId = lastRow[levelIdColumnIndex];
                            const selectedLevelId = selectRandomLevelId(currentLevelId, levelsData);
                            if (selectedLevelId !== null) {
                                finalLevelId = selectedLevelId.toString();
                                levelIdSource = 'random_from_levels';
                                console.log(`Sheet ${sheetName} 全新主题系列，从Levels文件随机选择Level_id: ${finalLevelId} (上一行值: ${currentLevelId})`);
                            } else {
                                // 随机选择失败，使用上一行数据
                                finalLevelId = lastRow[levelIdColumnIndex] || '1';
                                levelIdSource = 'previous_row_fallback';
                                console.log(`Sheet ${sheetName} 随机选择Level_id失败，使用上一行数据: ${finalLevelId}`);
                            }
                        } else {
                            // 无法加载Levels数据，使用上一行数据
                            finalLevelId = lastRow[levelIdColumnIndex] || '1';
                            levelIdSource = 'previous_row_default';
                            console.log(`Sheet ${sheetName} 全新主题系列，Levels数据不可用，使用上一行数据: ${finalLevelId}`);
                        }
                    }

                    newRow[levelIdColumnIndex] = finalLevelId;
                    console.log(`Sheet ${sheetName} 最终设置Level_id: ${levelIdColumnName} = ${finalLevelId} (来源: ${levelIdSource})`);
                } else {
                    console.warn(`Sheet ${sheetName} 中找不到Level_id相关列`);
                }

                // 处理UGC特定字段设置
                applyUGCFieldSettings(sheetName, headerRow, newRow);

                // 设置L列和M列的主题名称
                const lColumnIndex = 11; // L列通常是第12列（索引11）
                const mColumnIndex = 12; // M列通常是第13列（索引12）

                if (lColumnIndex < newRow.length) {
                    newRow[lColumnIndex] = themeName;
                    console.log(`Sheet ${sheetName} 设置L列(索引${lColumnIndex})主题名称: ${themeName}`);
                } else {
                    console.warn(`Sheet ${sheetName} L列索引${lColumnIndex}超出范围，当前行长度: ${newRow.length}`);
                }

                if (mColumnIndex < newRow.length) {
                    newRow[mColumnIndex] = themeName;
                    console.log(`Sheet ${sheetName} 设置M列(索引${mColumnIndex})主题名称: ${themeName}`);
                } else {
                    console.warn(`Sheet ${sheetName} M列索引${mColumnIndex}超出范围，当前行长度: ${newRow.length}`);
                }

                console.log(`Sheet ${sheetName} 新行:`, newRow);

                // 同系列主题排序插入功能（仅针对Custom_Ground_Color工作表）
                if (sheetName === 'Custom_Ground_Color' && smartConfig.similarity.isSimilar) {
                    console.log('=== 检测到同系列主题，开始排序插入处理 ===');

                    // 查找同系列最后主题的ID
                    const lastSimilarThemeId = findLastSimilarThemeId(smartConfig.similarity.matchedTheme, data, idColumnIndex);

                    if (lastSimilarThemeId) {
                        // 计算目标Level_show_id值
                        const targetLevelShowId = lastSimilarThemeId - 1;
                        console.log(`同系列最后主题ID: ${lastSimilarThemeId}, 目标Level_show_id: ${targetLevelShowId}`);

                        // 添加新行到数据（先添加，再进行排序插入）
                        data.push(newRow);

                        // 执行排序插入操作
                        const sortResult = performSortedInsertion(data, headerRow, newId, targetLevelShowId, themeName);

                        if (sortResult.success) {
                            console.log(`✅ 排序插入成功: 新主题 "${themeName}" 已插入到合适位置`);
                            console.log(`插入详情: 目标行=${sortResult.targetRowIndex}, 插入行=${sortResult.insertRowIndex}, 影响行数=${sortResult.movedRowsCount}`);
                        } else {
                            console.warn(`⚠️ 排序插入失败: ${sortResult.error}，使用默认添加方式`);
                        }
                    } else {
                        console.log('未找到同系列最后主题ID，使用默认添加方式');
                        // 添加新行到数据
                        data.push(newRow);
                    }
                } else {
                    // 非Custom_Ground_Color工作表或非同系列主题，使用默认添加方式
                    console.log(`Sheet ${sheetName}: 使用默认添加方式 (同系列: ${smartConfig.similarity.isSimilar})`);
                    // 添加新行到数据
                    data.push(newRow);
                }

                // 更新worksheet
                const newWorksheet = XLSX.utils.aoa_to_sheet(data);
                workbook.Sheets[sheetName] = newWorksheet;

                // 记录处理结果
                const sheetResult = {
                    sheetName: sheetName,
                    newId: newId,
                    newRowIndex: data.length - 1,
                    multiLangProcessed: levelNameColumnIndex !== -1,
                    multiLangColumn: levelNameColumnIndex !== -1 ? headerRow[levelNameColumnIndex] : null,
                    multiLangValue: levelNameColumnIndex !== -1 ? newRow[levelNameColumnIndex] : null,
                    levelIdProcessed: levelIdColumnIndex !== -1,
                    levelIdColumn: levelIdColumnIndex !== -1 ? headerRow[levelIdColumnIndex] : null,
                    levelIdValue: levelIdColumnIndex !== -1 ? newRow[levelIdColumnIndex] : null,
                    levelIdSource: levelIdColumnIndex !== -1 ? levelIdSource : null
                };

                processedSheets.push(sheetResult);
                console.log(`Sheet ${sheetName} 处理结果:`, sheetResult);
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

        // 将更新后的数据写回主工作表
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

        // 替换主工作表
        workbook.Sheets[originalSheetName] = newWorksheet;
        console.log(`✅ 主工作表 "${originalSheetName}" 已更新`);

        // 处理目标工作表（严格限制：仅限Light、ColorInfo）
        // 重要约束：不修改RSC_Theme.xls文件中的其他工作表，保持零影响原则
        console.log('=== 开始处理目标工作表 ===');
        const targetSheets = ['Light', 'ColorInfo'];
        if (rscAllSheetsData) {
            targetSheets.forEach(sheetName => {
                if (sheetName !== originalSheetName && rscAllSheetsData[sheetName]) {
                    console.log(`处理目标工作表: ${sheetName}`);

                    const sheetData = rscAllSheetsData[sheetName];
                    if (sheetData && sheetData.length > 0) {
                        // 创建新的工作表
                        const updatedWorksheet = XLSX.utils.aoa_to_sheet(sheetData);

                        // 保持原有的工作表属性
                        const originalSheet = workbook.Sheets[sheetName];
                        if (originalSheet && originalSheet['!ref']) {
                            updatedWorksheet['!ref'] = XLSX.utils.encode_range({
                                s: { c: 0, r: 0 },
                                e: {
                                    c: sheetData[0].length - 1,
                                    r: sheetData.length - 1
                                }
                            });
                        }

                        // 替换工作表
                        workbook.Sheets[sheetName] = updatedWorksheet;
                        console.log(`✅ 工作表 "${sheetName}" 已更新，数据行数: ${sheetData.length}`);

                        // 对于Light和ColorInfo工作表，输出详细的数据验证信息
                        if (sheetName === 'Light' || sheetName === 'ColorInfo') {
                            console.log(`${sheetName} 工作表数据验证:`);
                            if (sheetData.length > 1) {
                                const headerRow = sheetData[0];
                                const lastDataRow = sheetData[sheetData.length - 1];
                                console.log(`  表头: ${JSON.stringify(headerRow)}`);
                                console.log(`  最后一行数据: ${JSON.stringify(lastDataRow)}`);
                            }
                        }
                    } else {
                        console.warn(`工作表 "${sheetName}" 数据为空，跳过更新`);
                    }
                }
            });
        } else {
            console.warn('rscAllSheetsData 不存在，跳过其他工作表处理');
        }

        console.log('✅ 所有工作表处理完成');
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
     * 处理RSC_Language文件的多语言ID管理
     * @param {string} themeName - 主题名称
     * @param {number} multiLangId - 多语言ID
     */
    async function processRSCLanguage(themeName, multiLangId) {
        console.log('=== 开始处理RSC_Language文件 ===');
        console.log('主题名称:', themeName);
        console.log('多语言ID:', multiLangId);

        try {
            // 检查是否有RSC_Language文件句柄
            if (!folderManager || !folderManager.rscLanguageHandle) {
                console.log('RSC_Language.xls文件未找到，跳过多语言处理');
                return {
                    success: false,
                    skipped: true,
                    reason: 'RSC_Language.xls文件未找到，请确保该文件位于UGCTheme.xls同级目录'
                };
            }

            // 检查多语言配置
            if (!multiLangId || isNaN(multiLangId) || multiLangId <= 0) {
                console.log('多语言ID无效，跳过多语言处理');
                return {
                    success: false,
                    skipped: true,
                    reason: '多语言ID无效或未提供'
                };
            }

            // 读取RSC_Language文件
            console.log('读取RSC_Language.xls文件...');
            const languageFile = await folderManager.rscLanguageHandle.getFile();
            const languageArrayBuffer = await languageFile.arrayBuffer();
            const languageWorkbook = XLSX.read(languageArrayBuffer, { type: 'array' });

            // 查找rsc_Language工作表
            const languageSheetName = 'rsc_Language';
            if (!languageWorkbook.SheetNames.includes(languageSheetName)) {
                throw new Error(`RSC_Language.xls文件中未找到"${languageSheetName}"工作表`);
            }

            const languageSheet = languageWorkbook.Sheets[languageSheetName];
            const languageData = XLSX.utils.sheet_to_json(languageSheet, { header: 1 });

            console.log('RSC_Language文件读取成功，数据行数:', languageData.length);

            // 处理多语言数据
            const result = await updateLanguageData(languageData, multiLangId, themeName);

            if (result.updated) {
                // 保存更新后的文件
                const updatedSheet = XLSX.utils.aoa_to_sheet(result.data);
                languageWorkbook.Sheets[languageSheetName] = updatedSheet;

                await saveRSCLanguageFile(languageWorkbook);

                console.log('✅ RSC_Language文件处理完成');
                return {
                    success: true,
                    updated: true,
                    message: result.message
                };
            } else {
                console.log('RSC_Language文件无需更新');
                return {
                    success: true,
                    updated: false,
                    message: result.message
                };
            }

        } catch (error) {
            console.error('RSC_Language文件处理失败:', error);
            return {
                success: false,
                error: error.message
            };
        }
    }

    /**
     * 更新语言数据
     * @param {Array} languageData - 语言数据数组
     * @param {number} multiLangId - 多语言ID
     * @param {string} themeName - 主题名称
     */
    async function updateLanguageData(languageData, multiLangId, themeName) {
        console.log('=== 开始更新语言数据 ===');

        if (!languageData || languageData.length === 0) {
            throw new Error('语言数据为空');
        }

        // 查找表头行
        const headerRow = languageData[0];
        if (!headerRow) {
            throw new Error('未找到表头行');
        }

        // 查找列索引
        const idColumnIndex = headerRow.findIndex(col => col && col.toString().toLowerCase() === 'id');
        const notesColumnIndex = headerRow.findIndex(col => col && col.toString().toLowerCase() === 'notes');
        const chineseColumnIndex = headerRow.findIndex(col => col && col.toString().toLowerCase() === 'chinese');

        console.log('列索引:', { idColumnIndex, notesColumnIndex, chineseColumnIndex });

        if (idColumnIndex === -1 || notesColumnIndex === -1 || chineseColumnIndex === -1) {
            throw new Error('未找到必要的列：id、notes、chinese');
        }

        // 检查是否已存在该多语言ID
        let existingRowIndex = -1;
        for (let i = 1; i < languageData.length; i++) {
            const row = languageData[i];
            if (row && row[idColumnIndex] && parseInt(row[idColumnIndex]) === multiLangId) {
                existingRowIndex = i;
                break;
            }
        }

        if (existingRowIndex !== -1) {
            console.log(`多语言ID ${multiLangId} 已存在于第 ${existingRowIndex + 1} 行，跳过处理`);
            return {
                updated: false,
                message: `多语言ID ${multiLangId} 已存在，无需添加`
            };
        }

        // 处理主题名称：去除末尾数字
        const processedThemeName = themeName.replace(/\d+$/, '').trim();
        console.log('处理后的主题名称:', processedThemeName);

        // 在表格末尾添加新行
        const newRow = new Array(headerRow.length).fill('');
        newRow[idColumnIndex] = multiLangId;
        newRow[notesColumnIndex] = '主题';
        newRow[chineseColumnIndex] = processedThemeName;

        languageData.push(newRow);

        console.log(`✅ 已添加新的多语言记录: ID=${multiLangId}, 主题名称=${processedThemeName}`);

        return {
            updated: true,
            data: languageData,
            message: `已添加多语言ID ${multiLangId}，主题名称：${processedThemeName}`
        };
    }

    /**
     * 保存RSC_Language文件
     * @param {Object} workbook - 更新后的工作簿
     */
    async function saveRSCLanguageFile(workbook) {
        console.log('=== 开始保存RSC_Language文件 ===');

        try {
            if (!folderManager || !folderManager.rscLanguageHandle) {
                throw new Error('RSC_Language文件句柄不存在');
            }

            // 清理工作簿数据（类似于其他文件的处理）
            const cleanedWorkbook = createCleanWorkbook(workbook);

            // 生成Excel文件的二进制数据
            const excelBuffer = XLSX.write(cleanedWorkbook, {
                bookType: 'xls',
                type: 'array'
            });

            console.log('RSC_Language文件数据大小:', excelBuffer.length, 'bytes');

            // 验证当前权限
            const permission = await folderManager.rscLanguageHandle.queryPermission({ mode: 'readwrite' });
            console.log('RSC_Language文件当前权限:', permission);

            if (permission !== 'granted') {
                const newPermission = await folderManager.rscLanguageHandle.requestPermission({ mode: 'readwrite' });
                if (newPermission !== 'granted') {
                    throw new Error('无法获取RSC_Language文件写入权限');
                }
            }

            // 创建可写流并写入数据
            const writable = await folderManager.rscLanguageHandle.createWritable();
            await writable.write(excelBuffer);
            await writable.close();

            console.log('✅ RSC_Language文件保存成功');
            return true;

        } catch (error) {
            console.error('RSC_Language文件保存失败:', error);
            throw error;
        }
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
                        // 处理多语言文件（如果需要）
                        await handleMultiLanguageProcessing(themeName);

                        const message = ugcMessage ? `RSC_Theme文件保存成功，${ugcMessage}` : 'RSC_Theme文件已成功保存到原位置';
                        App.Utils.showStatus(message, 'success');

                        // 显示最终操作指引弹框
                        setTimeout(() => {
                            showFinalGuideModal();
                        }, 1000); // 延迟1秒显示，让用户看到成功消息

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

                                    // 显示最终操作指引弹框
                                    setTimeout(() => {
                                        showFinalGuideModal();
                                    }, 1000);
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
     * 处理多语言文件更新
     * @param {string} themeName - 主题名称
     */
    async function handleMultiLanguageProcessing(themeName) {
        console.log('=== 开始处理多语言文件更新 ===');

        try {
            // 获取多语言配置
            const multiLangConfig = getMultiLanguageConfig();
            console.log('多语言配置:', multiLangConfig);

            // 检查是否需要处理多语言
            if (!multiLangConfig.isValid) {
                console.log('多语言配置无效，跳过多语言处理');
                return;
            }

            // 处理RSC_Language文件
            const result = await processRSCLanguage(themeName, multiLangConfig.id);

            if (result.success) {
                if (result.updated) {
                    App.Utils.showStatus(`✅ 多语言文件已更新：${result.message}`, 'success', 5000);
                    console.log('多语言文件处理成功:', result.message);
                } else {
                    console.log('多语言文件无需更新:', result.message);
                }
            } else if (result.skipped) {
                console.log('多语言处理已跳过:', result.reason);
            } else {
                console.error('多语言文件处理失败:', result.error);
                App.Utils.showStatus(`⚠️ 多语言文件处理失败：${result.error}`, 'warning', 8000);
            }

        } catch (error) {
            console.error('多语言处理出错:', error);
            App.Utils.showStatus(`⚠️ 多语言处理出错：${error.message}`, 'warning', 8000);
        }
    }

    /**
     * 检查多语言处理的准备状态
     */
    function checkMultiLanguageReadiness() {
        const hasRSCLanguageFile = !!(folderManager && folderManager.rscLanguageHandle);
        const multiLangConfig = getMultiLanguageConfig();

        return {
            hasFile: hasRSCLanguageFile,
            hasValidConfig: multiLangConfig.isValid,
            config: multiLangConfig,
            ready: hasRSCLanguageFile && multiLangConfig.isValid
        };
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

            // 显示最终操作指引弹框（下载方式也需要手动操作）
            setTimeout(() => {
                showFinalGuideModal();
            }, 1000);

            // 下载完成后刷新数据预览
            console.log('文件下载完成，开始刷新数据预览...');
            refreshDataPreview();
        } catch (error) {
            console.error('文件下载失败:', error);
            App.Utils.showStatus('文件下载失败: ' + error.message, 'error');
        }
    }



    /**
     * 确认是否直接保存到原文件（简化版，仅保留直接保存选项）
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
                    <h3>💾 文件保存确认</h3>
                    <p><strong>主题数据处理完成，准备保存文件</strong></p>
                    <div style="margin: 20px 0;">
                        <button id="saveDirectBtn" style="background: #28a745; color: white; border: none; padding: 15px 30px; margin: 10px; border-radius: 8px; cursor: pointer; font-size: 16px;">
                            ✅ 直接保存到原文件
                        </button>
                    </div>
                    <p style="font-size: 14px; color: #666; margin-top: 15px;">
                        将直接覆盖原文件，无需手动替换<br>
                        <small>建议在保存前备份重要文件</small>
                    </p>
                </div>
            `;

            document.body.appendChild(modal);

            // 绑定事件
            document.getElementById('saveDirectBtn').onclick = () => {
                document.body.removeChild(modal);
                resolve(true);
            };

            // 点击背景关闭（默认选择直接保存）
            modal.addEventListener('click', (e) => {
                if (e.target === modal) {
                    document.body.removeChild(modal);
                    resolve(true);
                }
            });
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

        if (themeSelector) {
            themeSelector.value = '';
        }

        if (processThemeBtn) {
            processThemeBtn.disabled = true;
        }

        // 重置操作模式
        updateOperationMode('neutral');

        // 隐藏验证消息
        hideValidationMessage();

        // 隐藏相关区域
        const sections = ['themeInputSection', 'resultDisplay', 'dataPreview'];
        sections.forEach(sectionId => {
            const section = document.getElementById(sectionId);
            if (section) {
                section.style.display = 'none';
            }
        });

        // 隐藏配置面板
        toggleUGCConfigPanel(false);
        toggleLightConfigPanel(false);
        toggleColorInfoConfigPanel(false);
        toggleMultiLangPanel(false);

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

            // 显示选择中状态
            updateFileSelectionStatus('rscFileStatus', 'loading', '正在选择RSC_Theme文件...', '请在文件选择器中选择文件');

            // 选择RSC_Theme文件并获取写入权限
            const [fileHandle] = await window.showOpenFilePicker(pickerOptions);

            // 验证文件格式
            if (!fileHandle.name.toLowerCase().endsWith('.xls')) {
                updateFileSelectionStatus('rscFileStatus', 'error', '文件格式错误', '请选择.xls格式的RSC_Theme文件以确保Unity工具兼容性');
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
                updateFileSelectionStatus('rscFileStatus', 'error', '权限获取失败', '无法获取文件写入权限，请重新选择文件');
                App.Utils.showStatus('无法获取文件写入权限', 'error');
                return false;
            }

            // 读取文件内容
            const file = await fileHandle.getFile();

            // 显示加载状态
            updateFileSelectionStatus('rscFileStatus', 'loading', '正在加载文件...', `文件名: ${file.name}, 大小: ${formatFileSize(file.size)}`);
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

            // 显示成功状态
            const fileInfo = `文件名: ${file.name} | 大小: ${formatFileSize(file.size)} | 选择时间: ${getCurrentTimeString()}`;
            updateFileSelectionStatus('rscFileStatus', 'success', 'RSC_Theme文件选择成功', fileInfo);

            updateFileStatus('rscThemeStatus', `已加载 (支持直接保存): ${file.name}`, 'success');
            App.Utils.showStatus('RSC_Theme文件已加载，支持直接保存到原位置', 'success');
            checkReadyState();

            return true;
        } catch (error) {
            console.error('启用直接文件保存失败:', error);
            updateFileSelectionStatus('rscFileStatus', 'error', '文件加载失败', error.message);
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

            // 显示选择中状态
            updateFileSelectionStatus('ugcFileStatus', 'loading', '正在选择UGCTheme文件...', '请在文件选择器中选择文件');

            // 选择UGCTheme文件并获取写入权限
            const [fileHandle] = await window.showOpenFilePicker(pickerOptions);

            // 验证文件格式
            if (!fileHandle.name.toLowerCase().endsWith('.xls')) {
                updateFileSelectionStatus('ugcFileStatus', 'error', '文件格式错误', '请选择.xls格式的UGCTheme文件以确保Unity工具兼容性');
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
                updateFileSelectionStatus('ugcFileStatus', 'error', '权限获取失败', '无法获取文件写入权限，请重新选择文件');
                App.Utils.showStatus('无法获取文件写入权限', 'error');
                return false;
            }

            // 读取文件内容
            const file = await fileHandle.getFile();

            // 显示加载状态
            updateFileSelectionStatus('ugcFileStatus', 'loading', '正在加载文件...', `文件名: ${file.name}, 大小: ${formatFileSize(file.size)}`);
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

            // 显示成功状态
            const fileInfo = `文件名: ${file.name} | 大小: ${formatFileSize(file.size)} | 选择时间: ${getCurrentTimeString()}`;
            updateFileSelectionStatus('ugcFileStatus', 'success', 'UGCTheme文件选择成功', fileInfo);

            updateFileStatus('ugcThemeStatus', `已加载 (支持直接保存): ${file.name}`, 'success');
            App.Utils.showStatus('UGCTheme文件已加载，支持直接保存到原位置', 'success');
            checkReadyState();

            return true;
        } catch (error) {
            console.error('启用UGC直接文件保存失败:', error);
            updateFileSelectionStatus('ugcFileStatus', 'error', '文件加载失败', error.message);
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

            // 同步目标工作表（严格限制：仅限Light、ColorInfo）
            // 重要约束：不同步其他工作表，保持零影响原则
            const targetSheets = ['Light', 'ColorInfo'];
            targetSheets.forEach(sheetName => {
                if (sheetName !== mainSheetName && workbook.Sheets[sheetName]) {
                    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
                        header: 1,
                        defval: '',
                        raw: false
                    });

                    if (rscAllSheetsData) {
                        rscAllSheetsData[sheetName] = sheetData;
                        console.log(`已同步目标工作表 "${sheetName}" 数据`);
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
        refreshDataPreview: refreshDataPreview,

        // 多语言功能
        getMultiLanguageConfig: getMultiLanguageConfig,
        checkMultiLanguageReadiness: checkMultiLanguageReadiness,
        processRSCLanguage: processRSCLanguage,

        // 文件选择状态管理功能
        updateFileSelectionStatus: updateFileSelectionStatus,
        hideFileSelectionStatus: hideFileSelectionStatus,
        formatFileSize: formatFileSize,
        getCurrentTimeString: getCurrentTimeString
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

    /**
     * 初始化文件夹选择功能
     */
    function initializeFolderSelection() {
        console.log('初始化文件夹选择功能...');

        // 检查浏览器支持并设置动态UI
        const isSupported = App.UnityProjectFolderManager.isSupported();
        setupDynamicUI(isSupported);

        if (!isSupported) {
            console.log('浏览器不支持文件夹选择，已切换到兼容模式');
            return;
        }

        // 设置文件夹选择事件监听器
        if (selectFolderBtn) {
            selectFolderBtn.addEventListener('click', handleFolderSelection);
            console.log('文件夹选择按钮事件监听器已设置');
        }
    }

    /**
     * 根据浏览器能力设置动态UI显示
     */
    function setupDynamicUI(isSupported) {
        const folderSelectionSection = document.getElementById('folderSelectionSection');
        const individualFileSelectionSection = document.getElementById('individualFileSelectionSection');
        const browserUpgradeNote = document.getElementById('browserUpgradeNote');

        if (isSupported) {
            // 支持文件夹选择的浏览器：显示文件夹选择区域，隐藏单文件选择区域
            if (folderSelectionSection) {
                folderSelectionSection.style.display = 'block';
            }
            if (individualFileSelectionSection) {
                individualFileSelectionSection.style.display = 'none';
            }
            console.log('UI已切换到文件夹选择模式');
        } else {
            // 不支持文件夹选择的浏览器：隐藏文件夹选择区域，显示单文件选择区域
            if (folderSelectionSection) {
                folderSelectionSection.style.display = 'none';
            }
            if (individualFileSelectionSection) {
                individualFileSelectionSection.style.display = 'block';
            }
            if (browserUpgradeNote) {
                browserUpgradeNote.style.display = 'block';
            }
            console.log('UI已切换到兼容模式（单文件选择）');
        }
    }







    /**
     * 处理文件夹选择
     */
    async function handleFolderSelection() {
        try {
            console.log('开始文件夹选择...');
            App.Utils.showStatus('正在选择文件夹...', 'info');

            // 创建文件夹管理器实例
            folderManager = App.UnityProjectFolderManager.create();

            // 选择文件夹并自动定位文件
            const result = await folderManager.selectUnityProjectFolder();

            if (result.success) {
                console.log('文件夹选择成功:', result);
                await handleFolderSelectionSuccess(result);
            } else {
                throw new Error('文件夹选择失败');
            }

        } catch (error) {
            console.error('文件夹选择失败:', error);
            handleFolderSelectionError(error);
        }
    }

    /**
     * 处理文件夹选择成功
     */
    async function handleFolderSelectionSuccess(result) {
        try {
            console.log('处理文件夹选择成功结果...');

            // 设置文件夹选择模式标志
            folderSelectionActive = true;

            // 存储文件夹路径信息到 folderManager 实例
            if (folderManager && result.selectedFolderPath) {
                folderManager.selectedFolderPath = result.selectedFolderPath;
            }

            // 更新UI显示
            updateFolderSelectionUI(result);

            // 自动加载找到的文件
            let loadedCount = 0;

            if (result.rscThemeFound && result.files.rscTheme.hasPermission) {
                try {
                    const rscData = await folderManager.loadThemeFileData('rsc');
                    await setRSCThemeDataFromFolder(rscData);
                    loadedCount++;
                    console.log('RSC_Theme文件加载成功');
                } catch (error) {
                    console.error('RSC_Theme文件加载失败:', error);
                    updateFileStatusInUI('rsc', '加载失败', 'error');
                }
            }

            if (result.ugcThemeFound && result.files.ugcTheme.hasPermission) {
                try {
                    const ugcData = await folderManager.loadThemeFileData('ugc');
                    await setUGCThemeDataFromFolder(ugcData);
                    loadedCount++;
                    console.log('UGCTheme文件加载成功');
                } catch (error) {
                    console.error('UGCTheme文件加载失败:', error);
                    updateFileStatusInUI('ugc', '加载失败', 'error');
                }
            }

            // 加载RSC_Language文件（如果存在）
            if (result.rscLanguageFound && result.files.rscLanguage.hasPermission) {
                try {
                    console.log('开始加载RSC_Language文件...');
                    const languageFile = await folderManager.rscLanguageHandle.getFile();
                    const languageArrayBuffer = await languageFile.arrayBuffer();
                    const languageWorkbook = XLSX.read(languageArrayBuffer, { type: 'array' });

                    rscLanguageData = {
                        fileHandle: folderManager.rscLanguageHandle,
                        workbook: languageWorkbook,
                        fileName: languageFile.name,
                        size: languageFile.size
                    };

                    console.log('RSC_Language文件加载成功:', languageFile.name);
                    App.Utils.showStatus('RSC_Language文件已加载，支持多语言ID管理', 'info', 3000);
                } catch (error) {
                    console.error('RSC_Language文件加载失败:', error);
                    App.Utils.showStatus('RSC_Language文件加载失败，多语言功能将不可用', 'warning', 5000);
                }
            } else if (result.rscLanguageFound) {
                console.warn('RSC_Language.xls文件找到但权限获取失败');
                App.Utils.showStatus('RSC_Language文件权限获取失败，多语言功能将不可用', 'warning', 5000);
            } else {
                console.log('未找到RSC_Language.xls文件，多语言功能将不可用');
            }

            // 设置Levels文件信息（用于Level_id处理）
            if (result.levelsFound && result.files.levels.hasPermission) {
                // 创建unityProjectFiles对象（如果不存在）
                if (!unityProjectFiles) {
                    unityProjectFiles = {};
                }

                // 将Levels文件句柄添加到unityProjectFiles
                unityProjectFiles.levelsFile = result.files.levels.handle;
                console.log('Levels.xls文件信息已设置，可用于Level_id智能处理');
            } else if (result.levelsFound) {
                console.warn('Levels.xls文件找到但权限获取失败，Level_id将使用默认处理');
            } else {
                console.warn('未找到Levels.xls文件，Level_id将使用默认处理');
            }

            // 检查就绪状态
            checkReadyState();

            // 显示成功消息
            const message = `文件夹选择成功！已自动加载 ${loadedCount} 个主题文件`;
            App.Utils.showStatus(message, 'success');

        } catch (error) {
            console.error('处理文件夹选择结果失败:', error);
            App.Utils.showStatus('文件夹处理失败: ' + error.message, 'error');
        }
    }

    /**
     * 处理文件夹选择错误
     */
    function handleFolderSelectionError(error) {
        console.error('文件夹选择错误:', error);

        const errorHandlers = {
            'NotAllowedError': '用户拒绝了文件夹访问权限，请重新尝试并允许访问',
            'SecurityError': '安全限制：请确保在HTTPS环境下使用此功能',
            'AbortError': '用户取消了文件夹选择',
            'NotFoundError': '指定的文件夹不存在或已被移动',
            'InvalidStateError': '文件夹状态无效，请重新选择',
            'TypeError': '文件夹选择功能初始化失败，请刷新页面重试'
        };

        let message = errorHandlers[error.name] || `文件夹选择失败: ${error.message}`;

        // 为特定错误提供解决建议
        if (error.message.includes('Tools/xlsx')) {
            message += '\n\n💡 建议：请确保选择的是Unity项目中包含RSC_Theme.xls或UGCTheme.xls文件的文件夹';
        }

        App.Utils.showStatus(message, 'error');

        // 重置文件夹选择状态
        resetFolderSelection();

        // 显示降级选项提示
        showFallbackOptions();
    }

    /**
     * 重置文件夹选择状态
     */
    function resetFolderSelection() {
        folderSelectionActive = false;

        if (folderManager) {
            folderManager.cleanup();
            folderManager = null;
        }

        // 隐藏文件夹选择结果
        if (folderSelectionResult) {
            folderSelectionResult.style.display = 'none';
        }

        // 重置状态显示
        if (selectedFolderPath) {
            selectedFolderPath.textContent = '-';
            selectedFolderPath.style.color = '#666';
        }

        updateFileStatusInUI('rsc', '-', 'info');
        updateFileStatusInUI('ugc', '-', 'info');
    }

    /**
     * 显示降级选项
     */
    function showFallbackOptions() {
        const fallbackMessage = `
            <div style="margin-top: 15px; padding: 10px; border: 1px solid #17a2b8; border-radius: 5px; background-color: #d1ecf1;">
                <p style="margin: 0; color: #0c5460; font-weight: bold;">💡 替代方案</p>
                <p style="margin: 5px 0 0 0; color: #0c5460; font-size: 14px;">
                    您可以使用下方的"分别选择文件"功能，手动选择RSC_Theme.xls和UGCTheme.xls文件
                </p>
            </div>
        `;

        // 在文件夹选择区域显示降级提示
        if (folderUploadArea) {
            const existingFallback = folderUploadArea.querySelector('.fallback-message');
            if (existingFallback) {
                existingFallback.remove();
            }

            const fallbackDiv = document.createElement('div');
            fallbackDiv.className = 'fallback-message';
            fallbackDiv.innerHTML = fallbackMessage;
            folderUploadArea.appendChild(fallbackDiv);

            // 5秒后自动隐藏
            setTimeout(() => {
                if (fallbackDiv && fallbackDiv.parentNode) {
                    fallbackDiv.remove();
                }
            }, 5000);
        }
    }

    /**
     * 验证选择的文件夹
     */
    async function validateSelectedFolder(result) {
        const issues = [];

        if (!result.rscThemeFound) {
            issues.push('未找到RSC_Theme.xls文件');
        }

        if (!result.ugcThemeFound) {
            issues.push('未找到UGCTheme.xls文件');
        }

        if (issues.length > 0) {
            const message = `文件夹验证警告：\n${issues.join('\n')}\n\n是否继续使用找到的文件？`;
            return confirm(message);
        }

        return true;
    }

    /**
     * 更新文件夹选择UI显示
     */
    function updateFolderSelectionUI(result) {
        if (!folderSelectionResult) return;

        // 显示结果区域
        folderSelectionResult.style.display = 'block';

        // 更新文件夹路径
        if (selectedFolderPath) {
            selectedFolderPath.textContent = result.directoryPath;
            selectedFolderPath.style.color = '#28a745';
        }

        // 更新RSC文件状态
        updateFileStatusInUI('rsc',
            result.rscThemeFound ? '✅ 已找到并获取权限' : '❌ 未找到',
            result.rscThemeFound ? 'success' : 'error'
        );

        // 更新UGC文件状态
        updateFileStatusInUI('ugc',
            result.ugcThemeFound ? '✅ 已找到并获取权限' : '❌ 未找到',
            result.ugcThemeFound ? 'success' : 'error'
        );
    }

    /**
     * 更新单个文件状态显示
     */
    function updateFileStatusInUI(fileType, status, statusType) {
        const statusElement = fileType === 'rsc' ? rscFileStatus : ugcFileStatus;
        if (!statusElement) return;

        statusElement.textContent = status;

        // 设置状态颜色
        const colors = {
            'success': '#28a745',
            'error': '#dc3545',
            'warning': '#ffc107',
            'info': '#17a2b8'
        };

        statusElement.style.color = colors[statusType] || '#666';
    }

    /**
     * 从文件夹设置RSC主题数据
     */
    async function setRSCThemeDataFromFolder(fileData) {
        try {
            console.log('设置RSC主题数据（来自文件夹）...');

            // 设置RSC主题数据，包含文件句柄
            rscThemeData = {
                workbook: fileData.workbook,
                data: null, // 将在下面设置
                fileName: fileData.fileName,
                fileHandle: fileData.fileHandle, // 重要：保存文件句柄用于直接保存
                lastModified: fileData.lastModified
            };

            // 解析第一个工作表的数据
            const firstSheetName = fileData.workbook.SheetNames[0];
            const worksheet = fileData.workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                defval: '',
                raw: false
            });

            rscThemeData.data = jsonData;

            // 存储所有Sheet数据
            rscAllSheetsData = {};
            fileData.workbook.SheetNames.forEach(sheetName => {
                const sheet = fileData.workbook.Sheets[sheetName];
                const sheetData = XLSX.utils.sheet_to_json(sheet, {
                    header: 1,
                    defval: '',
                    raw: false
                });
                rscAllSheetsData[sheetName] = sheetData;
            });

            // 更新文件状态显示
            updateFileStatus('rscThemeStatus', `已加载 (文件夹模式): ${fileData.fileName}`, 'success');

            console.log('RSC主题数据设置完成（文件夹模式）');

        } catch (error) {
            console.error('设置RSC主题数据失败:', error);
            throw error;
        }
    }

    /**
     * 从文件夹设置UGC主题数据
     */
    async function setUGCThemeDataFromFolder(fileData) {
        try {
            console.log('设置UGC主题数据（来自文件夹）...');

            // 设置UGC主题数据，包含文件句柄
            ugcThemeData = {
                workbook: fileData.workbook,
                data: null, // 将在下面设置
                fileName: fileData.fileName,
                fileHandle: fileData.fileHandle, // 重要：保存文件句柄用于直接保存
                lastModified: fileData.lastModified
            };

            // 解析第一个工作表的数据
            const firstSheetName = fileData.workbook.SheetNames[0];
            const worksheet = fileData.workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                defval: '',
                raw: false
            });

            ugcThemeData.data = jsonData;

            // 存储所有Sheet数据
            ugcAllSheetsData = {};
            fileData.workbook.SheetNames.forEach(sheetName => {
                const sheet = fileData.workbook.Sheets[sheetName];
                const sheetData = XLSX.utils.sheet_to_json(sheet, {
                    header: 1,
                    defval: '',
                    raw: false
                });
                ugcAllSheetsData[sheetName] = sheetData;
            });

            // 更新文件状态显示
            updateFileStatus('ugcThemeStatus', `已加载 (文件夹模式): ${fileData.fileName}`, 'success');

            console.log('UGC主题数据设置完成（文件夹模式）');

        } catch (error) {
            console.error('设置UGC主题数据失败:', error);
            throw error;
        }
    }

    /**
     * 加载现有主题的Light配置
     */
    function loadExistingLightConfig(themeName) {
        console.log('加载现有主题的Light配置:', themeName);

        if (!rscAllSheetsData || !rscAllSheetsData['Light']) {
            console.log('RSC_Theme Light数据未加载，使用默认值');
            resetLightConfigToDefaults();
            return;
        }

        // 在RSC_Theme的Light sheet中查找主题
        const lightData = rscAllSheetsData['Light'];
        const lightHeaderRow = lightData[0];
        const lightNotesColumnIndex = lightHeaderRow.findIndex(col => col === 'notes');

        if (lightNotesColumnIndex === -1) {
            console.log('RSC_Theme Light sheet没有notes列，使用默认值');
            resetLightConfigToDefaults();
            return;
        }

        // 查找主题在Light中的行号
        const lightThemeRowIndex = lightData.findIndex((row, index) =>
            index > 0 && row[lightNotesColumnIndex] === themeName
        );

        if (lightThemeRowIndex === -1) {
            console.log(`在RSC_Theme Light sheet中未找到主题 "${themeName}"，使用默认值`);
            resetLightConfigToDefaults();
            return;
        }

        console.log(`在RSC_Theme Light sheet中找到主题 "${themeName}"，行索引: ${lightThemeRowIndex}`);

        // 加载Light配置值
        const lightRow = lightData[lightThemeRowIndex];
        const lightFieldMapping = {
            'Max': 'lightMax',
            'Dark': 'lightDark',
            'Min': 'lightMin',
            'SpecularLevel': 'lightSpecularLevel',
            'Gloss': 'lightGloss',
            'SpecularColor': 'lightSpecularColor'
        };

        Object.entries(lightFieldMapping).forEach(([lightColumn, inputId]) => {
            const columnIndex = lightHeaderRow.findIndex(col => col === lightColumn);
            if (columnIndex !== -1) {
                const value = lightRow[columnIndex] || '';
                const input = document.getElementById(inputId);
                if (input) {
                    input.value = value;

                    // 更新颜色预览
                    if (inputId === 'lightSpecularColor' && window.App && window.App.ColorPicker) {
                        window.App.ColorPicker.updateColorPreview(inputId, value);
                    }
                }
            }
        });
    }

    /**
     * 加载现有主题的ColorInfo配置
     */
    function loadExistingColorInfoConfig(themeName) {
        console.log('加载现有主题的ColorInfo配置:', themeName);

        if (!rscAllSheetsData || !rscAllSheetsData['ColorInfo']) {
            console.log('RSC_Theme ColorInfo数据未加载，使用默认值');
            resetColorInfoConfigToDefaults();
            return;
        }

        // 在RSC_Theme的ColorInfo sheet中查找主题
        const colorInfoData = rscAllSheetsData['ColorInfo'];
        const colorInfoHeaderRow = colorInfoData[0];
        const colorInfoNotesColumnIndex = colorInfoHeaderRow.findIndex(col => col === 'notes');

        if (colorInfoNotesColumnIndex === -1) {
            console.log('RSC_Theme ColorInfo sheet没有notes列，使用默认值');
            resetColorInfoConfigToDefaults();
            return;
        }

        // 查找主题在ColorInfo中的行号
        const colorInfoThemeRowIndex = colorInfoData.findIndex((row, index) =>
            index > 0 && row[colorInfoNotesColumnIndex] === themeName
        );

        if (colorInfoThemeRowIndex === -1) {
            console.log(`在RSC_Theme ColorInfo sheet中未找到主题 "${themeName}"，使用默认值`);
            resetColorInfoConfigToDefaults();
            return;
        }

        console.log(`在RSC_Theme ColorInfo sheet中找到主题 "${themeName}"，行索引: ${colorInfoThemeRowIndex}`);

        // 加载ColorInfo配置值
        const colorInfoRow = colorInfoData[colorInfoThemeRowIndex];
        const colorInfoFieldMapping = {
            'PickupDiffR': 'PickupDiffR',
            'PickupDiffG': 'PickupDiffG',
            'PickupDiffB': 'PickupDiffB',
            'PickupReflR': 'PickupReflR',
            'PickupReflG': 'PickupReflG',
            'PickupReflB': 'PickupReflB',
            'BallSpecR': 'BallSpecR',
            'BallSpecG': 'BallSpecG',
            'BallSpecB': 'BallSpecB',
            'ForegroundFogR': 'ForegroundFogR',
            'ForegroundFogG': 'ForegroundFogG',
            'ForegroundFogB': 'ForegroundFogB',
            'FogStart': 'FogStart',
            'FogEnd': 'FogEnd'
        };

        Object.entries(colorInfoFieldMapping).forEach(([columnName, inputId]) => {
            const columnIndex = colorInfoHeaderRow.findIndex(col => col === columnName);
            if (columnIndex !== -1) {
                const value = colorInfoRow[columnIndex];
                const input = document.getElementById(inputId);
                if (input && value !== undefined && value !== null && value !== '') {
                    input.value = value.toString();
                    console.log(`设置 ${inputId} = ${value}`);
                }
            }
        });

        // 更新颜色预览
        updateRgbColorPreview('PickupDiffR');
        updateRgbColorPreview('PickupReflR');
        updateRgbColorPreview('BallSpecR');
        updateRgbColorPreview('ForegroundFogR');

        console.log('ColorInfo配置加载完成');
    }

    /**
     * 更新颜色预览（辅助函数）
     */
    function updateColorPreview(inputId, color) {
        if (window.App && window.App.ColorPicker) {
            window.App.ColorPicker.updateColorPreview(inputId, color);
        }
    }



})();

// 模块加载完成日志
console.log('ThemeManager模块已加载 - 颜色主题管理功能已就绪');
