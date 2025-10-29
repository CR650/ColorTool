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
    let rscOriginalSheetsData = null; // 🔧 RSC_Theme文件的原始Sheet数据（用于重置非目标工作表）
    let ugcAllSheetsData = null;     // UGCTheme文件的所有Sheet数据
    let allObstacleData = null;      // AllObstacle.xls文件数据
    let multiLangConfig = null;      // 多语言配置数据
    let currentMappingMode = 'json'; // 当前映射模式：'json' 或 'direct'

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
     * 初始化主题管理模块（异步）
     */
    async function init() {
        if (isInitialized) {
            console.warn('ThemeManager模块已经初始化');
            return;
        }

        console.log('开始初始化ThemeManager模块...');

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

        // 异步加载对比映射数据
        console.log('开始加载映射数据...');
        await loadMappingData();
        console.log('映射数据加载完成');

        // 显示浏览器兼容性信息
        displayBrowserCompatibility();

        isInitialized = true;
        console.log('ThemeManager模块初始化完成');

        // 输出映射数据加载结果摘要
        if (mappingData) {
            const totalMappings = mappingData.data ? mappingData.data.length : 0;
            const validMappings = mappingData.data ? mappingData.data.filter(item =>
                item['RC现在的主题通道'] &&
                item['RC现在的主题通道'] !== '' &&
                item['颜色代码'] &&
                item['颜色代码'] !== ''
            ).length : 0;

            console.log(`📊 映射数据加载摘要: 总计${totalMappings}项映射，其中${validMappings}项有效`);

            if (validMappings > 17) {
                console.log(`✨ 检测到扩展映射数据！支持${validMappings}个颜色通道（超过默认的17个）`);
            }
        }
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

        // 初始化FloodLight配置验证
        initFloodLightValidation();

        // 初始化VolumetricFog配置验证
        initVolumetricFogValidation();
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
     * 初始化FloodLight配置验证
     */
    function initFloodLightValidation() {
        // TippingPoint字段验证 (0 到 5，支持一位小数)
        const tippingPointInput = document.getElementById('floodlightTippingPoint');
        if (tippingPointInput) {
            tippingPointInput.addEventListener('input', function() {
                validateFloodLightDecimalInput(this, 0, 5, 'TippingPoint');
            });
            tippingPointInput.addEventListener('blur', function() {
                validateFloodLightDecimalInput(this, 0, 5, 'TippingPoint');
            });
        }

        // Strength字段验证 (0 到 10，支持一位小数)
        const strengthInput = document.getElementById('floodlightStrength');
        if (strengthInput) {
            strengthInput.addEventListener('input', function() {
                validateFloodLightDecimalInput(this, 0, 10, 'Strength');
            });
            strengthInput.addEventListener('blur', function() {
                validateFloodLightDecimalInput(this, 0, 10, 'Strength');
            });
        }

        // 颜色字段验证
        const colorInput = document.getElementById('floodlightColor');
        if (colorInput) {
            colorInput.addEventListener('input', function() {
                validateColorInput(this);
                updateFloodLightColorPreview();
            });
            colorInput.addEventListener('blur', function() {
                validateColorInput(this);
                updateFloodLightColorPreview();
            });
        }

        // 初始化颜色预览
        updateFloodLightColorPreview();

        console.log('FloodLight配置验证已初始化');
    }

    /**
     * 初始化VolumetricFog配置验证
     */
    function initVolumetricFogValidation() {
        // Density字段验证 (0 到 20，支持一位小数)
        const densityInput = document.getElementById('volumetricfogDensity');
        if (densityInput) {
            densityInput.addEventListener('input', function() {
                validateVolumetricFogDecimalInput(this, 0, 20, 'Density');
            });
        }

        // X字段验证 (0 到 100，整数)
        const xInput = document.getElementById('volumetricfogX');
        if (xInput) {
            xInput.addEventListener('input', function() {
                validateVolumetricFogIntegerInput(this, 0, 100, 'X');
            });
        }

        // Y字段验证 (0 到 100，整数)
        const yInput = document.getElementById('volumetricfogY');
        if (yInput) {
            yInput.addEventListener('input', function() {
                validateVolumetricFogIntegerInput(this, 0, 100, 'Y');
            });
        }

        // Z字段验证 (0 到 100，整数)
        const zInput = document.getElementById('volumetricfogZ');
        if (zInput) {
            zInput.addEventListener('input', function() {
                validateVolumetricFogIntegerInput(this, 0, 100, 'Z');
            });
        }

        // Rotate字段验证 (-90 到 90，整数)
        const rotateInput = document.getElementById('volumetricfogRotate');
        if (rotateInput) {
            rotateInput.addEventListener('input', function() {
                validateVolumetricFogIntegerInput(this, -90, 90, 'Rotate');
            });
        }

        // 颜色字段验证
        const colorInput = document.getElementById('volumetricfogColor');
        if (colorInput) {
            colorInput.addEventListener('input', function() {
                updateVolumetricFogColorPreview();
            });
        }

        // 初始化颜色预览
        updateVolumetricFogColorPreview();

        console.log('VolumetricFog配置验证已初始化');
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
     * 验证FloodLight小数输入框 (支持一位小数)
     */
    function validateFloodLightDecimalInput(input, min, max, fieldName) {
        const value = parseFloat(input.value);

        // 移除之前的错误样式
        input.classList.remove('validation-error');

        // 移除之前的错误提示
        const existingError = input.parentElement.querySelector('.validation-message');
        if (existingError) {
            existingError.remove();
        }

        if (isNaN(value) || value < min || value > max) {
            // 添加错误样式
            input.classList.add('validation-error');

            // 创建错误提示
            const errorMsg = document.createElement('div');
            errorMsg.className = 'validation-message error';
            errorMsg.textContent = `${fieldName}范围: ${min}-${max}，支持一位小数`;
            input.parentElement.appendChild(errorMsg);

            // 自动修正值
            const correctedValue = Math.max(min, Math.min(max, isNaN(value) ? (min + max) / 2 : value));
            input.value = correctedValue.toFixed(1);

            console.warn(`${fieldName} 值无效: ${input.value}，已自动修正为: ${correctedValue.toFixed(1)}`);
        } else {
            // 确保值保持一位小数格式
            input.value = value.toFixed(1);
            console.log(`${fieldName} 值有效: ${value.toFixed(1)}`);
        }
    }

    /**
     * 更新FloodLight颜色预览
     */
    function updateFloodLightColorPreview() {
        const colorInput = document.getElementById('floodlightColor');
        const preview = document.getElementById('floodlightColorPreview');

        if (colorInput && preview) {
            const color = colorInput.value || 'FFFFFF';
            const cleanColor = color.replace('#', '');
            preview.style.backgroundColor = `#${cleanColor}`;
        }
    }

    /**
     * VolumetricFog小数输入验证
     * @param {HTMLInputElement} input - 输入元素
     * @param {number} min - 最小值
     * @param {number} max - 最大值
     * @param {string} fieldName - 字段名称
     */
    function validateVolumetricFogDecimalInput(input, min, max, fieldName) {
        let value = parseFloat(input.value);

        if (isNaN(value)) {
            console.warn(`VolumetricFog ${fieldName}字段输入无效，重置为默认值`);
            input.value = min === 0 ? '10.0' : '0.0';
            return;
        }

        // 限制小数位数为1位
        value = Math.round(value * 10) / 10;

        // 范围验证
        if (value < min) {
            console.warn(`VolumetricFog ${fieldName}字段值 ${value} 小于最小值 ${min}，自动修正`);
            value = min;
        } else if (value > max) {
            console.warn(`VolumetricFog ${fieldName}字段值 ${value} 大于最大值 ${max}，自动修正`);
            value = max;
        }

        input.value = value.toFixed(1);
    }

    /**
     * VolumetricFog整数输入验证
     * @param {HTMLInputElement} input - 输入元素
     * @param {number} min - 最小值
     * @param {number} max - 最大值
     * @param {string} fieldName - 字段名称
     */
    function validateVolumetricFogIntegerInput(input, min, max, fieldName) {
        let value = parseInt(input.value);

        if (isNaN(value)) {
            console.warn(`VolumetricFog ${fieldName}字段输入无效，重置为默认值`);
            input.value = min >= 0 ? min : 0;
            return;
        }

        // 范围验证
        if (value < min) {
            console.warn(`VolumetricFog ${fieldName}字段值 ${value} 小于最小值 ${min}，自动修正`);
            value = min;
        } else if (value > max) {
            console.warn(`VolumetricFog ${fieldName}字段值 ${value} 大于最大值 ${max}，自动修正`);
            value = max;
        }

        input.value = value;
    }

    /**
     * 更新VolumetricFog颜色预览
     */
    function updateVolumetricFogColorPreview() {
        const colorInput = document.getElementById('volumetricfogColor');
        const preview = document.getElementById('volumetricfogColorPreview');

        if (colorInput && preview) {
            const color = colorInput.value || 'FFFFFF';
            const cleanColor = color.replace('#', '');
            preview.style.backgroundColor = `#${cleanColor}`;
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
     * 显示/隐藏FloodLight配置面板
     */
    function toggleFloodLightConfigPanel(show) {
        const panel = document.getElementById('floodlightConfigSection');
        if (panel) {
            panel.style.display = show ? 'block' : 'none';
            console.log('FloodLight配置面板', show ? '已显示' : '已隐藏');
        }
    }

    /**
     * 显示/隐藏VolumetricFog配置面板
     */
    function toggleVolumetricFogConfigPanel(show) {
        const panel = document.getElementById('volumetricfogConfigSection');
        if (panel) {
            panel.style.display = show ? 'block' : 'none';
            console.log('VolumetricFog配置面板', show ? '已显示' : '已隐藏');
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
     * 获取表中第一个主题的Light配置数据
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

        if (lightNotesColumnIndex === -1 || lightData.length <= 5) {
            console.log('RSC_Theme Light sheet没有notes列或数据不足（需要至少6行），使用硬编码默认值');
            return {
                lightMax: '0',
                lightDark: '0',
                lightMin: '0',
                lightSpecularLevel: '100',
                lightGloss: '100',
                lightSpecularColor: 'FFFFFF'
            };
        }

        // 获取第一个主题数据（第6行，行索引为5，前5行是元数据）
        const firstRowIndex = 5;
        const firstRow = lightData[firstRowIndex];

        console.log(`读取表中第一个主题的Light配置，行索引: ${firstRowIndex}`);

        // 构建字段映射
        const lightFieldMapping = {
            'Max': 'lightMax',
            'Dark': 'lightDark',
            'Min': 'lightMin',
            'SpecularLevel': 'lightSpecularLevel',
            'Gloss': 'lightGloss',
            'SpecularColor': 'lightSpecularColor'
        };

        const firstThemeConfig = {};
        Object.entries(lightFieldMapping).forEach(([columnName, fieldId]) => {
            const columnIndex = lightHeaderRow.findIndex(col => col === columnName);
            if (columnIndex !== -1) {
                const value = firstRow[columnIndex];
                firstThemeConfig[fieldId] = (value !== undefined && value !== null && value !== '') ? value.toString() : '0';
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
                firstThemeConfig[fieldId] = defaults[fieldId] || '0';
            }
        });

        console.log('第一个主题的Light配置:', firstThemeConfig);
        return firstThemeConfig;
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
     * 获取表中第一个主题的FloodLight配置数据
     */
    function getLastThemeFloodLightConfig() {
        if (!rscAllSheetsData || !rscAllSheetsData['FloodLight']) {
            console.log('RSC_Theme FloodLight数据未加载，使用硬编码默认值');
            return {
                floodlightColor: 'FFFFFF',
                floodlightTippingPoint: '2.5',
                floodlightStrength: '7.8',
                floodlightIsOn: false,
                floodlightJumpActiveIsLightOn: false
            };
        }

        const floodLightData = rscAllSheetsData['FloodLight'];
        const floodLightHeaderRow = floodLightData[0];
        const floodLightNotesColumnIndex = floodLightHeaderRow.findIndex(col => col === 'notes');

        if (floodLightNotesColumnIndex === -1 || floodLightData.length <= 5) {
            console.log('RSC_Theme FloodLight sheet没有notes列或数据不足（需要至少6行），使用硬编码默认值');
            return {
                floodlightColor: 'FFFFFF',
                floodlightTippingPoint: '2.5',
                floodlightStrength: '7.8',
                floodlightIsOn: false,
                floodlightJumpActiveIsLightOn: false
            };
        }

        // FloodLight字段映射
        const floodLightFieldMapping = {
            'Color': 'floodlightColor',
            'TippingPoint': 'floodlightTippingPoint',
            'Strength': 'floodlightStrength',
            'IsOn': 'floodlightIsOn',
            'JumpActiveIsLightOn': 'floodlightJumpActiveIsLightOn'
        };

        // 获取第一个主题数据（第6行，行索引为5，前5行是元数据）
        const firstRow = floodLightData[5];

        const firstThemeConfig = {};
        Object.entries(floodLightFieldMapping).forEach(([columnName, fieldId]) => {
            const columnIndex = floodLightHeaderRow.findIndex(col => col === columnName);
            if (columnIndex !== -1) {
                const value = firstRow[columnIndex];
                if (fieldId === 'floodlightTippingPoint' || fieldId === 'floodlightStrength') {
                    // 将存储的整数值转换为小数显示（除以10）
                    const numValue = parseInt(value) || 0;
                    firstThemeConfig[fieldId] = (numValue / 10).toFixed(1);
                } else if (fieldId === 'floodlightIsOn' || fieldId === 'floodlightJumpActiveIsLightOn') {
                    // 转换为布尔值
                    firstThemeConfig[fieldId] = value === 1 || value === '1' || value === true;
                } else {
                    // 颜色值
                    firstThemeConfig[fieldId] = (value !== undefined && value !== null && value !== '') ? value.toString() : 'FFFFFF';
                }
            } else {
                // 如果找不到列，使用默认值
                const defaults = {
                    'floodlightColor': 'FFFFFF',
                    'floodlightTippingPoint': '2.5',
                    'floodlightStrength': '7.8',
                    'floodlightIsOn': false,
                    'floodlightJumpActiveIsLightOn': false
                };
                firstThemeConfig[fieldId] = defaults[fieldId];
            }
        });

        console.log('第一个主题的FloodLight配置:', firstThemeConfig);
        return firstThemeConfig;
    }

    /**
     * 获取表中第一个主题的VolumetricFog配置数据
     */
    function getLastThemeVolumetricFogConfig() {
        if (!rscAllSheetsData || !rscAllSheetsData['VolumetricFog']) {
            console.log('RSC_Theme VolumetricFog数据未加载，使用硬编码默认值');
            return {
                volumetricfogColor: 'FFFFFF',
                volumetricfogX: '50',
                volumetricfogY: '50',
                volumetricfogZ: '50',
                volumetricfogDensity: '10.0',
                volumetricfogRotate: '0',
                volumetricfogIsOn: false
            };
        }

        const volumetricFogData = rscAllSheetsData['VolumetricFog'];
        if (volumetricFogData.length <= 5) {
            console.log('VolumetricFog表数据不足（需要至少6行），使用硬编码默认值');
            return {
                volumetricfogColor: 'FFFFFF',
                volumetricfogX: '50',
                volumetricfogY: '50',
                volumetricfogZ: '50',
                volumetricfogDensity: '10.0',
                volumetricfogRotate: '0',
                volumetricfogIsOn: false
            };
        }

        const headerRow = volumetricFogData[0];
        const firstRowIndex = 5;
        const firstRow = volumetricFogData[firstRowIndex];

        console.log('VolumetricFog表第一个主题数据（第6行）:', firstRow);

        const firstThemeConfig = {};

        // VolumetricFog字段映射
        const volumetricFogFieldMapping = {
            'Color': 'volumetricfogColor',
            'X': 'volumetricfogX',
            'Y': 'volumetricfogY',
            'Z': 'volumetricfogZ',
            'Density': 'volumetricfogDensity',
            'Rotate': 'volumetricfogRotate',
            'IsOn': 'volumetricfogIsOn'
        };

        Object.entries(volumetricFogFieldMapping).forEach(([columnName, fieldId]) => {
            const columnIndex = headerRow.findIndex(col => col === columnName);
            if (columnIndex !== -1 && firstRow[columnIndex] !== undefined && firstRow[columnIndex] !== '') {
                let value = firstRow[columnIndex];

                // 特殊处理：Density字段需要÷10显示
                if (columnName === 'Density') {
                    value = (parseFloat(value) / 10).toFixed(1);
                } else if (columnName === 'IsOn') {
                    // 布尔值处理
                    value = value === 1 || value === '1' || value === true;
                }

                firstThemeConfig[fieldId] = value;
                console.log(`VolumetricFog ${columnName} -> ${fieldId}: ${value}`);
            } else {
                // 如果找不到列，使用默认值
                const defaults = {
                    'volumetricfogColor': 'FFFFFF',
                    'volumetricfogX': '50',
                    'volumetricfogY': '50',
                    'volumetricfogZ': '50',
                    'volumetricfogDensity': '10.0',
                    'volumetricfogRotate': '0',
                    'volumetricfogIsOn': false
                };
                firstThemeConfig[fieldId] = defaults[fieldId];
            }
        });

        console.log('第一个主题的VolumetricFog配置:', firstThemeConfig);
        return firstThemeConfig;
    }

    /**
     * 获取表中第一个主题的ColorInfo配置数据
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

        if (colorInfoNotesColumnIndex === -1 || colorInfoData.length <= 5) {
            console.log('RSC_Theme ColorInfo sheet没有notes列或数据不足（需要至少6行），使用硬编码默认值');
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

        // 获取第一个主题数据（第6行，行索引为5，前5行是元数据）
        const firstRowIndex = 5;
        const firstRow = colorInfoData[firstRowIndex];

        console.log(`读取表中第一个主题的ColorInfo配置，行索引: ${firstRowIndex}`);

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

        const firstThemeConfig = {};
        Object.entries(colorInfoFieldMapping).forEach(([columnName, fieldId]) => {
            const columnIndex = colorInfoHeaderRow.findIndex(col => col === columnName);
            if (columnIndex !== -1) {
                const value = firstRow[columnIndex];
                firstThemeConfig[fieldId] = (value !== undefined && value !== null && value !== '') ? value.toString() : '0';
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
                firstThemeConfig[fieldId] = defaults[fieldId] || '0';
            }
        });

        console.log('第一个主题的ColorInfo配置:', firstThemeConfig);
        return firstThemeConfig;
    }

    /**
     * 重置FloodLight配置为默认值（新建主题时使用表中最后一个主题的数据）
     */
    function resetFloodLightConfigToDefaults() {
        const floodLightDefaults = getLastThemeFloodLightConfig();

        Object.entries(floodLightDefaults).forEach(([fieldId, defaultValue]) => {
            const input = document.getElementById(fieldId);
            if (input) {
                if (fieldId === 'floodlightIsOn' || fieldId === 'floodlightJumpActiveIsLightOn') {
                    // 处理checkbox
                    input.checked = defaultValue;
                } else {
                    // 处理普通输入框
                    input.value = defaultValue;

                    // 更新颜色预览
                    if (fieldId === 'floodlightColor') {
                        updateFloodLightColorPreview();
                    }
                }
            }
        });

        console.log('FloodLight配置已重置为最后一个主题的配置');
    }

    /**
     * 重置VolumetricFog配置为默认值（新建主题时使用表中最后一个主题的数据）
     */
    function resetVolumetricFogConfigToDefaults() {
        const volumetricFogDefaults = getLastThemeVolumetricFogConfig();

        Object.entries(volumetricFogDefaults).forEach(([fieldId, defaultValue]) => {
            const input = document.getElementById(fieldId);
            if (input) {
                if (input.type === 'checkbox') {
                    input.checked = defaultValue === 1 || defaultValue === true;
                } else {
                    input.value = defaultValue;
                }
                input.classList.remove('validation-error');

                // 移除错误提示
                const errorMsg = input.parentElement.querySelector('.validation-message');
                if (errorMsg) {
                    errorMsg.remove();
                }

                // 更新颜色预览
                if (fieldId === 'volumetricfogColor') {
                    updateVolumetricFogColorPreview();
                }
            }
        });

        console.log('VolumetricFog配置已重置为最后一个主题的配置');
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
            // 🔧 调试日志：检查为什么返回默认值
            if (value !== undefined && value !== null && value !== '') {
                console.warn(`⚠️ validateRgbValue: 输入值"${value}"无效，返回默认值${defaultValue}`);
            }
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
     * 验证FloodLight小数值（支持一位小数）
     */
    function validateFloodLightDecimal(value, defaultValue, min, max) {
        const numValue = parseFloat(value);
        if (isNaN(numValue) || numValue < min || numValue > max) {
            return defaultValue;
        }
        return Math.round(numValue * 10) / 10; // 保持一位小数
    }

    /**
     * 验证VolumetricFog整数值
     */
    function validateVolumetricFogInteger(value, defaultValue, min, max) {
        const numValue = parseInt(value);
        if (isNaN(numValue) || numValue < min || numValue > max) {
            return defaultValue;
        }
        return numValue;
    }

    /**
     * 验证VolumetricFog小数值（支持一位小数）
     */
    function validateVolumetricFogDecimal(value, defaultValue, min, max) {
        const numValue = parseFloat(value);
        if (isNaN(numValue) || numValue < min || numValue > max) {
            return defaultValue;
        }
        return Math.round(numValue * 10) / 10; // 保持一位小数
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
     * 获取FloodLight配置数据
     */
    function getFloodLightConfigData() {
        return {
            Color: validateHexColor(document.getElementById('floodlightColor')?.value, 'FFFFFF'),
            TippingPoint: Math.round(validateFloodLightDecimal(document.getElementById('floodlightTippingPoint')?.value, 2.5, 0, 5) * 10),
            Strength: Math.round(validateFloodLightDecimal(document.getElementById('floodlightStrength')?.value, 7.8, 0, 10) * 10),
            IsOn: document.getElementById('floodlightIsOn')?.checked ? 1 : 0,
            JumpActiveIsLightOn: document.getElementById('floodlightJumpActiveIsLightOn')?.checked ? 1 : 0,
            LightStrength: 180  // 固定值
        };
    }

    /**
     * 获取VolumetricFog配置数据
     */
    function getVolumetricFogConfigData() {
        return {
            Color: validateHexColor(document.getElementById('volumetricfogColor')?.value, 'FFFFFF'),
            X: validateVolumetricFogInteger(document.getElementById('volumetricfogX')?.value, 50, 0, 100),
            Y: validateVolumetricFogInteger(document.getElementById('volumetricfogY')?.value, 50, 0, 100),
            Z: validateVolumetricFogInteger(document.getElementById('volumetricfogZ')?.value, 50, 0, 100),
            Density: Math.round(validateVolumetricFogDecimal(document.getElementById('volumetricfogDensity')?.value, 10.0, 0, 20) * 10),
            Rotate: validateVolumetricFogInteger(document.getElementById('volumetricfogRotate')?.value, 0, -90, 90),
            IsOn: document.getElementById('volumetricfogIsOn')?.checked ? 1 : 0
        };
    }

    /**
     * 获取ColorInfo配置数据
     */
    function getColorInfoConfigData() {
        // 🔧 调试日志：检查UI元素的值
        const pickupDiffRElement = document.getElementById('PickupDiffR');
        const pickupDiffRValue = pickupDiffRElement?.value;
        console.log(`🔍 getColorInfoConfigData - PickupDiffR元素值: "${pickupDiffRValue}" (类型: ${typeof pickupDiffRValue})`);

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
     * @param {string} themeName - 主题名称
     * @param {boolean} isNewTheme - 是否为新建主题（默认false）
     */
    function loadExistingUGCConfig(themeName, isNewTheme = false) {
        console.log('\n=== 开始加载UGC配置（支持条件读取） ===');
        console.log('主题名称:', themeName);
        console.log('是否新建主题:', isNewTheme);
        console.log('ugcAllSheetsData状态:', ugcAllSheetsData ? '已加载' : '未加载');
        console.log('rscAllSheetsData状态:', rscAllSheetsData ? '已加载' : '未加载');
        console.log('sourceData状态:', sourceData ? '已加载' : '未加载');
        console.log('当前映射模式:', currentMappingMode);

        if (ugcAllSheetsData) {
            console.log('UGC数据包含的sheets:', Object.keys(ugcAllSheetsData));
        }

        if (!themeName) {
            console.log('主题名称为空，使用默认值');
            resetUGCConfigToDefaults();
            return;
        }

        // 检查是否为直接映射模式
        const isDirectMode = currentMappingMode === 'direct';
        console.log(`是否为直接映射模式: ${isDirectMode}`);

        // 🔧 新建主题模式下，如果是直接映射模式且有源数据，直接从源数据读取
        if (isNewTheme && isDirectMode && sourceData && sourceData.workbook) {
            console.log('🔧 新建主题（直接映射模式）：直接从源数据读取UGC配置');

            // 定义条件读取函数映射
            const conditionalReadFunctions = {
                'Custom_Ground_Color': findCustomGroundColorValueDirect,
                'Custom_Fragile_Color': findCustomFragileColorValueDirect,
                'Custom_Fragile_Active_Color': findCustomFragileActiveColorValueDirect,
                'Custom_Jump_Color': findCustomJumpColorValueDirect,
                'Custom_Jump_Active_Color': findCustomJumpActiveColorValueDirect
            };

            // 定义字段映射
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

            // 遍历每个sheet，使用条件读取函数从源数据读取
            Object.entries(sheetFieldMapping).forEach(([sheetName, mapping]) => {
                console.log(`\n--- 新建主题：从源数据加载Sheet: ${sheetName} 的UI配置 ---`);

                const conditionalReadFunc = conditionalReadFunctions[sheetName];
                if (!conditionalReadFunc) {
                    console.warn(`未找到Sheet ${sheetName} 的条件读取函数`);
                    return;
                }

                // 提取字段值
                Object.entries(mapping).forEach(([fieldKey, fieldValue]) => {
                    if (fieldKey.endsWith('Field')) {
                        const columnName = mapping[fieldKey.replace('Field', 'Column')];

                        // 🔧 使用条件读取函数，传递isNewTheme=true
                        const directValue = conditionalReadFunc(columnName, true, themeName);

                        if (directValue !== null && directValue !== undefined && directValue !== '') {
                            const finalValue = parseInt(directValue) || 0;
                            configData[fieldValue] = finalValue;
                            console.log(`✅ [源数据读取] Sheet ${sheetName} 字段 ${columnName}: ${finalValue} -> UI字段 ${fieldValue}`);
                        } else {
                            // 使用默认值
                            const defaultValue = fieldValue.includes('Alpha') ? 50 : 0;
                            configData[fieldValue] = defaultValue;
                            console.log(`⚠️ [使用默认值] Sheet ${sheetName} 字段 ${columnName}: ${defaultValue} -> UI字段 ${fieldValue}`);
                        }
                    }
                });
            });

            console.log('\n✅ 新建主题：最终加载的UGC配置数据（将显示在UI中）:', configData);
            setUGCConfigData(configData);
            return;
        }

        // 🔧 更新现有主题模式：需要从RSC_Theme和UGCTheme中查找主题
        if (!ugcAllSheetsData || !rscAllSheetsData) {
            console.log('UGC数据或RSC数据未加载，使用默认值');
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

            // 定义条件读取函数映射
            const conditionalReadFunctions = {
                'Custom_Ground_Color': findCustomGroundColorValueDirect,
                'Custom_Fragile_Color': findCustomFragileColorValueDirect,
                'Custom_Fragile_Active_Color': findCustomFragileActiveColorValueDirect,
                'Custom_Jump_Color': findCustomJumpColorValueDirect,
                'Custom_Jump_Active_Color': findCustomJumpActiveColorValueDirect
            };

            const configData = {};

            // 遍历每个sheet查找主题数据（使用行号匹配）
            Object.entries(sheetFieldMapping).forEach(([sheetName, mapping]) => {
                console.log(`\n--- 加载Sheet: ${sheetName} 的UI配置 ---`);

                // 提取字段值（支持条件读取）
                Object.entries(mapping).forEach(([fieldKey, fieldValue]) => {
                    if (fieldKey.endsWith('Field')) {
                        const columnName = mapping[fieldKey.replace('Field', 'Column')];
                        let finalValue = 0;

                        // 如果是直接映射模式，尝试使用条件读取函数从源数据读取
                        if (isDirectMode && sourceData && conditionalReadFunctions[sheetName]) {
                            const conditionalReadFunc = conditionalReadFunctions[sheetName];
                            const directValue = conditionalReadFunc(columnName, false, themeName);

                            if (directValue !== null && directValue !== undefined && directValue !== '') {
                                finalValue = parseInt(directValue) || 0;
                                console.log(`✅ [源数据读取] Sheet ${sheetName} 字段 ${columnName}: ${finalValue} -> UI字段 ${fieldValue}`);
                            } else {
                                // 条件读取返回空，从UGCTheme文件读取
                                const sheetData = ugcAllSheetsData[sheetName];
                                if (sheetData && sheetData.length > targetRowNumber) {
                                    const headerRow = sheetData[0];
                                    const themeRow = sheetData[targetRowNumber];
                                    const columnIndex = headerRow.findIndex(col => col === columnName);

                                    if (columnIndex !== -1) {
                                        const value = themeRow[columnIndex];
                                        finalValue = value !== undefined && value !== '' ? parseInt(value) || 0 : 0;
                                        console.log(`📋 [UGCTheme文件] Sheet ${sheetName} 字段 ${columnName}: ${finalValue} -> UI字段 ${fieldValue}`);
                                    } else {
                                        console.log(`⚠️ Sheet ${sheetName} 未找到列 ${columnName}，使用默认值0`);
                                    }
                                } else {
                                    console.log(`⚠️ Sheet ${sheetName} 数据不足，使用默认值0`);
                                }
                            }
                        } else {
                            // 非直接映射模式，从UGCTheme文件读取
                            const sheetData = ugcAllSheetsData[sheetName];
                            if (sheetData && sheetData.length > targetRowNumber) {
                                const headerRow = sheetData[0];
                                const themeRow = sheetData[targetRowNumber];
                                const columnIndex = headerRow.findIndex(col => col === columnName);

                                if (columnIndex !== -1) {
                                    const value = themeRow[columnIndex];
                                    finalValue = value !== undefined && value !== '' ? parseInt(value) || 0 : 0;
                                    console.log(`📋 [UGCTheme文件] Sheet ${sheetName} 字段 ${columnName}: ${finalValue} -> UI字段 ${fieldValue}`);
                                } else {
                                    console.log(`⚠️ Sheet ${sheetName} 未找到列 ${columnName}，使用默认值0`);
                                }
                            } else {
                                console.log(`⚠️ Sheet ${sheetName} 数据不足，使用默认值0`);
                            }
                        }

                        configData[fieldValue] = finalValue;
                    }
                });
            });

            console.log('\n✅ 最终加载的UGC配置数据（将显示在UI中）:', configData);
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
     * 从JSON文件异步读取映射数据
     * @returns {Promise<Object|null>} 映射数据对象或null（如果读取失败）
     */
    async function fetchMappingDataFromJSON() {
        // 方法1：尝试使用fetch API（适用于HTTP服务器环境）
        try {
            console.log('尝试从XLS/对比.json文件读取映射数据（方法1：fetch API）...');

            const response = await fetch('XLS/对比.json');
            if (!response.ok) {
                throw new Error(`HTTP错误: ${response.status} ${response.statusText}`);
            }

            const jsonData = await response.json();
            console.log('JSON文件读取成功，开始验证数据格式...');

            // 验证JSON数据格式
            if (!validateMappingData(jsonData)) {
                throw new Error('JSON文件格式验证失败');
            }

            console.log('✅ JSON映射数据验证通过，包含', jsonData.data.length, '个映射项');
            return jsonData;

        } catch (fetchError) {
            console.warn('⚠️ fetch方法失败:', fetchError.message);

            // 方法2：尝试使用XMLHttpRequest（可能在某些环境下工作）
            try {
                console.log('尝试备用方法（XMLHttpRequest）...');
                return await fetchMappingDataWithXHR();
            } catch (xhrError) {
                console.warn('⚠️ XMLHttpRequest方法也失败:', xhrError.message);

                // 方法3：检查是否在file://协议下运行，给出明确提示
                if (window.location.protocol === 'file:') {
                    console.warn('🚨 检测到file://协议，无法读取JSON映射文件');
                    console.warn('💡 解决方案：');
                    console.warn('   方案1: 双击项目根目录的 start_server.bat 文件');
                    console.warn('   方案2: 手动运行 python -m http.server 8000');
                    console.warn('   方案3: 使用 Live Server 等开发工具');
                    console.warn('   然后访问: http://localhost:8000');

                    // 显示用户友好的提示
                    if (App.Utils) {
                        App.Utils.showStatus(
                            '⚠️ 检测到本地文件访问限制。请使用HTTP服务器访问项目以获得完整功能。',
                            'warning',
                            8000
                        );
                    }
                }

                return null;
            }
        }
    }

    /**
     * 使用XMLHttpRequest读取映射数据（备用方法）
     * @returns {Promise<Object>} 映射数据对象
     */
    function fetchMappingDataWithXHR() {
        return new Promise((resolve, reject) => {
            const xhr = new XMLHttpRequest();
            xhr.open('GET', 'XLS/对比.json', true);
            xhr.responseType = 'json';

            xhr.onload = function() {
                if (xhr.status === 200 || xhr.status === 0) { // status 0 for file:// protocol
                    try {
                        let jsonData = xhr.response;

                        // 如果响应不是对象，尝试解析
                        if (typeof jsonData === 'string') {
                            jsonData = JSON.parse(jsonData);
                        }

                        if (!validateMappingData(jsonData)) {
                            reject(new Error('JSON文件格式验证失败'));
                            return;
                        }

                        console.log('✅ XMLHttpRequest方法成功，JSON映射数据验证通过');
                        resolve(jsonData);
                    } catch (parseError) {
                        reject(new Error('JSON解析失败: ' + parseError.message));
                    }
                } else {
                    reject(new Error(`XMLHttpRequest失败: ${xhr.status} ${xhr.statusText}`));
                }
            };

            xhr.onerror = function() {
                reject(new Error('XMLHttpRequest网络错误'));
            };

            xhr.send();
        });
    }

    /**
     * 验证映射数据格式
     * @param {Object} jsonData - 待验证的JSON数据
     * @returns {boolean} 验证是否通过
     */
    function validateMappingData(jsonData) {
        try {
            // 检查基本结构
            if (!jsonData || typeof jsonData !== 'object') {
                console.error('JSON数据不是有效对象');
                return false;
            }

            if (!Array.isArray(jsonData.data)) {
                console.error('JSON数据缺少data数组');
                return false;
            }

            // 检查必要字段
            const requiredFields = ['RC现在的主题通道', '颜色代码'];
            let validItemCount = 0;

            for (const item of jsonData.data) {
                if (!item || typeof item !== 'object') continue;

                const hasRequiredFields = requiredFields.every(field =>
                    item.hasOwnProperty(field)
                );

                if (hasRequiredFields) {
                    validItemCount++;
                }
            }

            if (validItemCount === 0) {
                console.error('JSON数据中没有包含必要字段的有效映射项');
                return false;
            }

            console.log(`JSON数据验证通过: 总计${jsonData.data.length}项，有效映射${validItemCount}项`);
            return true;

        } catch (error) {
            console.error('JSON数据验证过程中出错:', error);
            return false;
        }
    }

    /**
     * 获取内置映射数据（向后兼容）
     * @returns {Object} 内置映射数据对象
     */
    function getBuiltinMappingData() {
        return {
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
                }
                // 注意：移除了P3→P3, P4→P4, P7→P7, P8→P8等直接映射
                // 这些通道应该只通过JSON文件中的明确映射关系来处理
                // 例如：P11→P3, P15→P4等间接映射是正确的
                // 但P3→P3, P4→P4等直接映射会导致不必要的颜色查找
            ]
        };
    }

    /**
     * 加载对比映射数据（支持动态JSON文件读取）
     */
    async function loadMappingData() {
        try {
            console.log('=== 开始加载映射数据 ===');

            // 首先尝试从JSON文件读取
            const jsonMappingData = await fetchMappingDataFromJSON();

            if (jsonMappingData) {
                // JSON文件读取成功
                mappingData = jsonMappingData;
                updateFileStatus('mappingStatus', '已加载 (JSON文件)', 'success');
                console.log('✅ 对比映射数据加载成功（来源：XLS/对比.json文件）');
                console.log('映射数据统计:', {
                    总项目数: mappingData.data.length,
                    有效通道数: mappingData.data.filter(item =>
                        item['RC现在的主题通道'] &&
                        item['RC现在的主题通道'] !== '' &&
                        item['颜色代码'] &&
                        item['颜色代码'] !== ''
                    ).length
                });
            } else {
                // JSON文件读取失败，使用内置数据
                console.log('🔄 回退到内置映射数据...');
                mappingData = getBuiltinMappingData();
                updateFileStatus('mappingStatus', '已加载 (内置数据)', 'warning');
                console.log('⚠️ 对比映射数据加载成功（来源：内置数据，功能受限）');
            }

        } catch (error) {
            console.error('❌ 映射数据加载过程中出现错误:', error);

            // 出现异常时使用内置数据
            mappingData = getBuiltinMappingData();
            updateFileStatus('mappingStatus', '已加载 (内置数据)', 'error');
            console.log('🔄 因错误回退到内置映射数据');
        }

        console.log('=== 映射数据加载完成 ===');
    }

    /**
     * 检测映射模式
     * @param {Object} sourceData - 源数据对象
     * @returns {string} 映射模式：'direct' 或 'json'
     */
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
                const statusData = XLSX.utils.sheet_to_json(statusSheet, {
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

    /**
     * 解析Status工作表，提取Color字段状态
     * @param {Object} sourceData - 源数据对象
     * @returns {Object} Status状态信息对象
     */
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

            const statusData = XLSX.utils.sheet_to_json(statusSheet, {
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

            let colorStatus = 0;
            let hasColorField = false;
            if (colorColumnIndex !== -1) {
                const colorStatusValue = statusRow[colorColumnIndex];
                colorStatus = parseInt(colorStatusValue) || 0;
                hasColorField = true;
                console.log(`Color字段状态: ${colorStatus} (原始值: "${colorStatusValue}")`);
            } else {
                console.log('Status工作表中未找到Color列');
            }

            // 查找Light列的索引
            const lightColumnIndex = headers.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim().toUpperCase();
                return headerStr === 'LIGHT';
            });

            let lightStatus = 0;
            let hasLightField = false;
            if (lightColumnIndex !== -1) {
                const lightStatusValue = statusRow[lightColumnIndex];
                lightStatus = parseInt(lightStatusValue) || 0;
                hasLightField = true;
                console.log(`Light字段状态: ${lightStatus} (原始值: "${lightStatusValue}")`);
            } else {
                console.log('Status工作表中未找到Light列');
            }

            // 查找ColorInfo列的索引
            const colorInfoColumnIndex = headers.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim().toUpperCase();
                return headerStr === 'COLORINFO';
            });

            let colorInfoStatus = 0;
            let hasColorInfoField = false;
            if (colorInfoColumnIndex !== -1) {
                const colorInfoStatusValue = statusRow[colorInfoColumnIndex];
                colorInfoStatus = parseInt(colorInfoStatusValue) || 0;
                hasColorInfoField = true;
                console.log(`ColorInfo字段状态: ${colorInfoStatus} (原始值: "${colorInfoStatusValue}")`);
            } else {
                console.log('Status工作表中未找到ColorInfo列');
            }

            // 查找VolumetricFog列的索引
            const volumetricFogColumnIndex = headers.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim().toUpperCase();
                return headerStr === 'VOLUMETRICFOG';
            });

            let volumetricFogStatus = 0;
            let hasVolumetricFogField = false;
            if (volumetricFogColumnIndex !== -1) {
                const volumetricFogStatusValue = statusRow[volumetricFogColumnIndex];
                volumetricFogStatus = parseInt(volumetricFogStatusValue) || 0;
                hasVolumetricFogField = true;
                console.log(`VolumetricFog字段状态: ${volumetricFogStatus} (原始值: "${volumetricFogStatusValue}")`);
            } else {
                console.log('Status工作表中未找到VolumetricFog列');
            }

            // 查找FloodLight列的索引
            const floodLightColumnIndex = headers.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim().toUpperCase();
                return headerStr === 'FLOODLIGHT';
            });

            let floodLightStatus = 0;
            let hasFloodLightField = false;
            if (floodLightColumnIndex !== -1) {
                const floodLightStatusValue = statusRow[floodLightColumnIndex];
                floodLightStatus = parseInt(floodLightStatusValue) || 0;
                hasFloodLightField = true;
                console.log(`FloodLight字段状态: ${floodLightStatus} (原始值: "${floodLightStatusValue}")`);
            } else {
                console.log('Status工作表中未找到FloodLight列');
            }

            // 查找Custom_Ground_Color列的索引
            const customGroundColorColumnIndex = headers.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim().toUpperCase();
                return headerStr === 'CUSTOM_GROUND_COLOR';
            });

            let customGroundColorStatus = 0;
            let hasCustomGroundColorField = false;
            if (customGroundColorColumnIndex !== -1) {
                const customGroundColorStatusValue = statusRow[customGroundColorColumnIndex];
                customGroundColorStatus = parseInt(customGroundColorStatusValue) || 0;
                hasCustomGroundColorField = true;
                console.log(`Custom_Ground_Color字段状态: ${customGroundColorStatus} (原始值: "${customGroundColorStatusValue}")`);
            } else {
                console.log('Status工作表中未找到Custom_Ground_Color列');
            }

            // 查找Custom_Fragile_Color列的索引
            const customFragileColorColumnIndex = headers.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim().toUpperCase();
                return headerStr === 'CUSTOM_FRAGILE_COLOR';
            });

            let customFragileColorStatus = 0;
            let hasCustomFragileColorField = false;
            if (customFragileColorColumnIndex !== -1) {
                const customFragileColorStatusValue = statusRow[customFragileColorColumnIndex];
                customFragileColorStatus = parseInt(customFragileColorStatusValue) || 0;
                hasCustomFragileColorField = true;
                console.log(`Custom_Fragile_Color字段状态: ${customFragileColorStatus} (原始值: "${customFragileColorStatusValue}")`);
            } else {
                console.log('Status工作表中未找到Custom_Fragile_Color列');
            }

            // 查找Custom_Fragile_Active_Color列的索引
            const customFragileActiveColorColumnIndex = headers.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim().toUpperCase();
                return headerStr === 'CUSTOM_FRAGILE_ACTIVE_COLOR';
            });

            let customFragileActiveColorStatus = 0;
            let hasCustomFragileActiveColorField = false;
            if (customFragileActiveColorColumnIndex !== -1) {
                const customFragileActiveColorStatusValue = statusRow[customFragileActiveColorColumnIndex];
                customFragileActiveColorStatus = parseInt(customFragileActiveColorStatusValue) || 0;
                hasCustomFragileActiveColorField = true;
                console.log(`Custom_Fragile_Active_Color字段状态: ${customFragileActiveColorStatus} (原始值: "${customFragileActiveColorStatusValue}")`);
            } else {
                console.log('Status工作表中未找到Custom_Fragile_Active_Color列');
            }

            // 查找Custom_Jump_Color列的索引
            const customJumpColorColumnIndex = headers.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim().toUpperCase();
                return headerStr === 'CUSTOM_JUMP_COLOR';
            });

            let customJumpColorStatus = 0;
            let hasCustomJumpColorField = false;
            if (customJumpColorColumnIndex !== -1) {
                const customJumpColorStatusValue = statusRow[customJumpColorColumnIndex];
                customJumpColorStatus = parseInt(customJumpColorStatusValue) || 0;
                hasCustomJumpColorField = true;
                console.log(`Custom_Jump_Color字段状态: ${customJumpColorStatus} (原始值: "${customJumpColorStatusValue}")`);
            } else {
                console.log('Status工作表中未找到Custom_Jump_Color列');
            }

            // 查找Custom_Jump_Active_Color列的索引
            const customJumpActiveColorColumnIndex = headers.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim().toUpperCase();
                return headerStr === 'CUSTOM_JUMP_ACTIVE_COLOR';
            });

            let customJumpActiveColorStatus = 0;
            let hasCustomJumpActiveColorField = false;
            if (customJumpActiveColorColumnIndex !== -1) {
                const customJumpActiveColorStatusValue = statusRow[customJumpActiveColorColumnIndex];
                customJumpActiveColorStatus = parseInt(customJumpActiveColorStatusValue) || 0;
                hasCustomJumpActiveColorField = true;
                console.log(`Custom_Jump_Active_Color字段状态: ${customJumpActiveColorStatus} (原始值: "${customJumpActiveColorStatusValue}")`);
            } else {
                console.log('Status工作表中未找到Custom_Jump_Active_Color列');
            }

            const result = {
                colorStatus: colorStatus,
                hasColorField: hasColorField,
                colorColumnIndex: colorColumnIndex,
                lightStatus: lightStatus,
                hasLightField: hasLightField,
                lightColumnIndex: lightColumnIndex,
                colorInfoStatus: colorInfoStatus,
                hasColorInfoField: hasColorInfoField,
                colorInfoColumnIndex: colorInfoColumnIndex,
                volumetricFogStatus: volumetricFogStatus,
                hasVolumetricFogField: hasVolumetricFogField,
                volumetricFogColumnIndex: volumetricFogColumnIndex,
                floodLightStatus: floodLightStatus,
                hasFloodLightField: hasFloodLightField,
                floodLightColumnIndex: floodLightColumnIndex,
                customGroundColorStatus: customGroundColorStatus,
                hasCustomGroundColorField: hasCustomGroundColorField,
                customGroundColorColumnIndex: customGroundColorColumnIndex,
                customFragileColorStatus: customFragileColorStatus,
                hasCustomFragileColorField: hasCustomFragileColorField,
                customFragileColorColumnIndex: customFragileColorColumnIndex,
                customFragileActiveColorStatus: customFragileActiveColorStatus,
                hasCustomFragileActiveColorField: hasCustomFragileActiveColorField,
                customFragileActiveColorColumnIndex: customFragileActiveColorColumnIndex,
                customJumpColorStatus: customJumpColorStatus,
                hasCustomJumpColorField: hasCustomJumpColorField,
                customJumpColorColumnIndex: customJumpColorColumnIndex,
                customJumpActiveColorStatus: customJumpActiveColorStatus,
                hasCustomJumpActiveColorField: hasCustomJumpActiveColorField,
                customJumpActiveColorColumnIndex: customJumpActiveColorColumnIndex,
                headers: headers,
                statusRow: statusRow,
                isColorValid: colorStatus === 1,
                isLightValid: lightStatus === 1,
                isColorInfoValid: colorInfoStatus === 1,
                isVolumetricFogValid: volumetricFogStatus === 1,
                isFloodLightValid: floodLightStatus === 1,
                isCustomGroundColorValid: customGroundColorStatus === 1,
                isCustomFragileColorValid: customFragileColorStatus === 1,
                isCustomFragileActiveColorValid: customFragileActiveColorStatus === 1,
                isCustomJumpColorValid: customJumpColorStatus === 1,
                isCustomJumpActiveColorValid: customJumpActiveColorStatus === 1
            };

            console.log('Status工作表解析结果:', result);
            return result;

        } catch (error) {
            console.error('解析Status工作表时出错:', error);
            return { colorStatus: 0, hasColorField: false, error: error.message };
        }
    }

    /**
     * 更新映射模式指示器
     * @param {string} mode - 映射模式
     * @param {Object} additionalInfo - 附加信息
     */
    function updateMappingModeIndicator(mode, additionalInfo = {}) {
        currentMappingMode = mode;

        // 更新源数据文件选择结果中的映射模式信息
        const sourceFileResult = document.getElementById('sourceFileSelectionResult');
        if (sourceFileResult && sourceFileResult.style.display !== 'none') {
            let mappingModeInfo = sourceFileResult.querySelector('.mapping-mode-info');

            if (!mappingModeInfo) {
                mappingModeInfo = document.createElement('div');
                mappingModeInfo.className = 'mapping-mode-info';
                mappingModeInfo.style.cssText = 'margin-top: 10px; padding: 8px; border-radius: 3px; font-size: 13px; font-weight: bold;';
                sourceFileResult.appendChild(mappingModeInfo);
            }

            if (mode === 'direct') {
                mappingModeInfo.style.backgroundColor = '#e7f3ff';
                mappingModeInfo.style.color = '#0066cc';
                mappingModeInfo.style.border = '1px solid #b3d9ff';
                const colorStatus = additionalInfo.colorStatus !== undefined ?
                    (additionalInfo.colorStatus === 1 ? '有效' : '无效') : '未知';
                mappingModeInfo.innerHTML = `
                    🎯 <strong>直接映射模式</strong><br>
                    <small>检测到Status工作表，Color状态: ${colorStatus}（${additionalInfo.fieldCount || 0}个字段）</small>
                `;
            } else {
                mappingModeInfo.style.backgroundColor = '#fff3cd';
                mappingModeInfo.style.color = '#856404';
                mappingModeInfo.style.border = '1px solid #ffeaa7';
                mappingModeInfo.innerHTML = `
                    📋 <strong>JSON映射模式</strong><br>
                    <small>使用XLS/对比.json文件进行映射关系处理</small>
                `;
            }
        }

        // 更新独立的映射模式指示器
        const mappingModeIndicator = document.getElementById('mappingModeIndicator');
        const mappingModeContent = document.getElementById('mappingModeContent');

        if (mappingModeIndicator && mappingModeContent) {
            // 显示指示器
            mappingModeIndicator.style.display = 'block';

            // 移除之前的模式类
            mappingModeIndicator.classList.remove('direct-mode', 'json-mode');

            if (mode === 'direct') {
                mappingModeIndicator.classList.add('direct-mode');
                const colorStatus = additionalInfo.colorStatus !== undefined ?
                    (additionalInfo.colorStatus === 1 ? '有效' : '无效') : '未知';
                const lightStatus = additionalInfo.lightStatus !== undefined ?
                    (additionalInfo.lightStatus === 1 ? '有效' : '无效') : '未知';
                const colorInfoStatus = additionalInfo.colorInfoStatus !== undefined ?
                    (additionalInfo.colorInfoStatus === 1 ? '有效' : '无效') : '未知';
                const volumetricFogStatus = additionalInfo.volumetricFogStatus !== undefined ?
                    (additionalInfo.volumetricFogStatus === 1 ? '有效' : '无效') : '未知';
                const floodLightStatus = additionalInfo.floodLightStatus !== undefined ?
                    (additionalInfo.floodLightStatus === 1 ? '有效' : '无效') : '未知';
                const customGroundColorStatus = additionalInfo.customGroundColorStatus !== undefined ?
                    (additionalInfo.customGroundColorStatus === 1 ? '有效' : '无效') : '未知';
                const customFragileColorStatus = additionalInfo.customFragileColorStatus !== undefined ?
                    (additionalInfo.customFragileColorStatus === 1 ? '有效' : '无效') : '未知';
                const customFragileActiveColorStatus = additionalInfo.customFragileActiveColorStatus !== undefined ?
                    (additionalInfo.customFragileActiveColorStatus === 1 ? '有效' : '无效') : '未知';
                const customJumpColorStatus = additionalInfo.customJumpColorStatus !== undefined ?
                    (additionalInfo.customJumpColorStatus === 1 ? '有效' : '无效') : '未知';
                const customJumpActiveColorStatus = additionalInfo.customJumpActiveColorStatus !== undefined ?
                    (additionalInfo.customJumpActiveColorStatus === 1 ? '有效' : '无效') : '未知';
                mappingModeContent.innerHTML = `
                    <div class="mapping-mode-title">
                        <span class="mapping-mode-icon">🎯</span>直接映射模式
                    </div>
                    <div class="mapping-mode-description">
                        <strong>RSC工作表状态:</strong> Color: ${colorStatus}, Light: ${lightStatus}, ColorInfo: ${colorInfoStatus}, VolumetricFog: ${volumetricFogStatus}, FloodLight: ${floodLightStatus}<br>
                        <strong>UGC工作表状态:</strong> Ground: ${customGroundColorStatus}, Fragile: ${customFragileColorStatus}, FragileActive: ${customFragileActiveColorStatus}, Jump: ${customJumpColorStatus}, JumpActive: ${customJumpActiveColorStatus}<br>
                        支持${additionalInfo.fieldCount || 0}个直接字段映射
                    </div>
                `;
            } else {
                mappingModeIndicator.classList.add('json-mode');
                mappingModeContent.innerHTML = `
                    <div class="mapping-mode-title">
                        <span class="mapping-mode-icon">📋</span>JSON映射模式
                    </div>
                    <div class="mapping-mode-description">
                        使用XLS/对比.json文件进行映射关系处理
                    </div>
                `;
            }
        }

        console.log(`映射模式已设置为: ${mode}`);
    }

    /**
     * 设置源数据
     * @param {Object} data - 源数据文件内容
     */
    function setSourceData(data) {
        sourceData = data;

        // 检测映射模式
        const detectedMode = detectMappingMode(data);

        // 获取附加信息用于显示
        let additionalInfo = {};
        if (detectedMode === 'direct' && data.workbook && data.workbook.Sheets['Status']) {
            try {
                // 解析Status工作表获取Color、Light、ColorInfo、VolumetricFog和FloodLight状态
                const statusInfo = parseStatusSheet(data);
                additionalInfo.colorStatus = statusInfo.colorStatus;
                additionalInfo.hasColorField = statusInfo.hasColorField;
                additionalInfo.lightStatus = statusInfo.lightStatus;
                additionalInfo.hasLightField = statusInfo.hasLightField;
                additionalInfo.colorInfoStatus = statusInfo.colorInfoStatus;
                additionalInfo.hasColorInfoField = statusInfo.hasColorInfoField;
                additionalInfo.volumetricFogStatus = statusInfo.volumetricFogStatus;
                additionalInfo.hasVolumetricFogField = statusInfo.hasVolumetricFogField;
                additionalInfo.floodLightStatus = statusInfo.floodLightStatus;
                additionalInfo.hasFloodLightField = statusInfo.hasFloodLightField;
                additionalInfo.customGroundColorStatus = statusInfo.customGroundColorStatus;
                additionalInfo.hasCustomGroundColorField = statusInfo.hasCustomGroundColorField;
                additionalInfo.customFragileColorStatus = statusInfo.customFragileColorStatus;
                additionalInfo.hasCustomFragileColorField = statusInfo.hasCustomFragileColorField;
                additionalInfo.customFragileActiveColorStatus = statusInfo.customFragileActiveColorStatus;
                additionalInfo.hasCustomFragileActiveColorField = statusInfo.hasCustomFragileActiveColorField;
                additionalInfo.customJumpColorStatus = statusInfo.customJumpColorStatus;
                additionalInfo.hasCustomJumpColorField = statusInfo.hasCustomJumpColorField;
                additionalInfo.customJumpActiveColorStatus = statusInfo.customJumpActiveColorStatus;
                additionalInfo.hasCustomJumpActiveColorField = statusInfo.hasCustomJumpActiveColorField;

                // 如果有Color工作表，计算字段数量
                if (data.workbook.Sheets['Color']) {
                    const colorData = XLSX.utils.sheet_to_json(data.workbook.Sheets['Color'], { header: 1 });
                    const headers = colorData[0] || [];
                    const directFields = [];
                    for (let i = 1; i <= 7; i++) directFields.push(`G${i}`);
                    for (let i = 1; i <= 49; i++) directFields.push(`P${i}`);
                    const foundFields = headers.filter(h => directFields.includes(h?.toString().trim().toUpperCase()));
                    additionalInfo.fieldCount = foundFields.length;
                } else {
                    additionalInfo.fieldCount = 0;
                }
            } catch (error) {
                console.warn('获取直接映射信息时出错:', error);
                additionalInfo.colorStatus = 0;
                additionalInfo.hasColorField = false;
                additionalInfo.fieldCount = 0;
            }
        }

        // 更新映射模式指示器
        updateMappingModeIndicator(detectedMode, additionalInfo);

        // 使用新的文件选择状态更新函数，保持与其他文件选择的一致性
        const fileInfo = `文件名: ${data.fileName} | 大小: ${formatFileSize(data.fileSize || 0)} | 选择时间: ${getCurrentTimeString()}`;
        updateFileSelectionStatus('sourceFileStatus', 'success', '源数据文件选择成功', fileInfo);

        // 调试：输出源数据结构
        console.log('源数据加载完成:', {
            fileName: data.fileName,
            headers: data.headers,
            dataCount: data.data.length,
            sampleData: data.data.slice(0, 3),
            mappingMode: detectedMode
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
                rscOriginalSheetsData = {}; // 🔧 同时保存原始数据的深拷贝
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const sheetData = XLSX.utils.sheet_to_json(worksheet, {
                        header: 1,
                        defval: '',
                        raw: false
                    });
                    rscAllSheetsData[sheetName] = sheetData;

                    // 🔧 深拷贝原始数据（用于后续重置非目标工作表）
                    rscOriginalSheetsData[sheetName] = JSON.parse(JSON.stringify(sheetData));
                });

                console.log('RSC_Theme所有Sheet数据已存储:', Object.keys(rscAllSheetsData));
                console.log('🔧 RSC_Theme原始Sheet数据已备份:', Object.keys(rscOriginalSheetsData));

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

            // 显示FloodLight配置面板并加载现有值
            toggleFloodLightConfigPanel(true);
            loadExistingFloodLightConfig(selectedTheme);

            // 显示VolumetricFog配置面板并加载现有值
            toggleVolumetricFogConfigPanel(true);
            loadExistingVolumetricFogConfig(selectedTheme);

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
            // 🔧 新建主题时，如果是直接映射模式，从源数据加载UGC配置
            if (currentMappingMode === 'direct' && sourceData && sourceData.workbook) {
                console.log('🔧 新建主题（直接映射模式）：尝试从源数据加载UGC配置到UI');
                loadExistingUGCConfig(inputValue, true); // 🔧 传递isNewTheme=true
            } else {
                // 非直接映射模式或无源数据，使用默认值
                resetUGCConfigToDefaults();
            }

            // 显示Light配置面板（新建主题时总是显示）
            toggleLightConfigPanel(true);
            // 🔧 新建主题时，如果是直接映射模式，从源数据加载Light配置
            if (currentMappingMode === 'direct' && sourceData && sourceData.workbook) {
                console.log('🔧 新建主题（直接映射模式）：尝试从源数据加载Light配置到UI');
                loadExistingLightConfig(inputValue, true); // 🔧 传递isNewTheme=true
            } else {
                // 非直接映射模式或无源数据，使用最后一个主题的Light配置作为默认值
                resetLightConfigToDefaults();
            }

            // 显示ColorInfo配置面板（新建主题时总是显示）
            toggleColorInfoConfigPanel(true);
            // 🔧 新建主题时，如果是直接映射模式，从源数据加载ColorInfo配置
            if (currentMappingMode === 'direct' && sourceData && sourceData.workbook) {
                console.log('🔧 新建主题（直接映射模式）：尝试从源数据加载ColorInfo配置到UI');
                loadExistingColorInfoConfig(inputValue, true); // 🔧 传递isNewTheme=true
            } else {
                // 非直接映射模式或无源数据，使用最后一个主题的ColorInfo配置作为默认值
                resetColorInfoConfigToDefaults();
            }

            // 显示FloodLight配置面板（新建主题时总是显示）
            toggleFloodLightConfigPanel(true);
            // 🔧 新建主题时，如果是直接映射模式，从源数据加载FloodLight配置
            if (currentMappingMode === 'direct' && sourceData && sourceData.workbook) {
                console.log('🔧 新建主题（直接映射模式）：尝试从源数据加载FloodLight配置到UI');
                loadExistingFloodLightConfig(inputValue, true); // 🔧 传递isNewTheme=true
            } else {
                // 非直接映射模式或无源数据，使用最后一个主题的FloodLight配置作为默认值
                resetFloodLightConfigToDefaults();
            }

            // 显示VolumetricFog配置面板（新建主题时总是显示）
            toggleVolumetricFogConfigPanel(true);
            // 🔧 新建主题时，如果是直接映射模式，从源数据加载VolumetricFog配置
            if (currentMappingMode === 'direct' && sourceData && sourceData.workbook) {
                console.log('🔧 新建主题（直接映射模式）：尝试从源数据加载VolumetricFog配置到UI');
                loadExistingVolumetricFogConfig(inputValue, true); // 🔧 传递isNewTheme=true
            } else {
                // 非直接映射模式或无源数据，使用最后一个主题的VolumetricFog配置作为默认值
                resetVolumetricFogConfigToDefaults();
            }

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

            // 隐藏FloodLight配置面板
            toggleFloodLightConfigPanel(false);

            // 隐藏VolumetricFog配置面板
            toggleVolumetricFogConfigPanel(false);

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

                // 处理AllObstacle文件（仅全新系列主题时）
                let allObstacleResult = null;
                if (result.isNewTheme) {
                    console.log('=== AllObstacle处理检查 ===');
                    console.log('检查是否需要处理AllObstacle文件...');
                    console.log('主题名称:', themeName);
                    console.log('是否为新增主题:', result.isNewTheme);

                    const smartConfig = getSmartMultiLanguageConfig(themeName);
                    console.log('智能配置检测结果:', {
                        isNewSeries: smartConfig.isNewSeries,
                        hasMultiLangConfig: !!smartConfig.multiLangConfig,
                        isMultiLangValid: smartConfig.multiLangConfig ? smartConfig.multiLangConfig.isValid : false,
                        multiLangId: smartConfig.multiLangConfig ? smartConfig.multiLangConfig.id : null
                    });

                    if (smartConfig.isNewSeries && smartConfig.multiLangConfig && smartConfig.multiLangConfig.isValid) {
                        console.log('✅ 满足AllObstacle处理条件，开始处理...');
                        console.log('多语言ID:', smartConfig.multiLangConfig.id);

                        allObstacleResult = await processAllObstacle(themeName, smartConfig.multiLangConfig.id);

                        if (allObstacleResult.success) {
                            console.log('✅ AllObstacle文件处理成功');
                            if (allObstacleResult.updated) {
                                console.log('📝 AllObstacle文件已更新:', allObstacleResult.message);
                            } else {
                                console.log('ℹ️ AllObstacle文件无需更新:', allObstacleResult.message);
                            }
                        } else if (allObstacleResult.skipped) {
                            console.log('⚠️ AllObstacle文件处理被跳过:', allObstacleResult.reason);
                        } else {
                            console.error('❌ AllObstacle文件处理失败:', allObstacleResult.error);
                            App.Utils.showStatus('AllObstacle文件处理失败: ' + allObstacleResult.error, 'warning', 5000);
                        }
                    } else {
                        console.log('❌ 不满足AllObstacle处理条件:');
                        if (!smartConfig.isNewSeries) {
                            console.log('  - 非全新系列主题');
                        }
                        if (!smartConfig.multiLangConfig) {
                            console.log('  - 多语言配置缺失');
                        }
                        if (smartConfig.multiLangConfig && !smartConfig.multiLangConfig.isValid) {
                            console.log('  - 多语言配置无效');
                        }
                        console.log('跳过AllObstacle处理');
                    }
                } else {
                    console.log('=== AllObstacle处理检查 ===');
                    console.log('非新增主题，跳过AllObstacle处理');
                }
                console.log('=== AllObstacle处理检查完成 ===');

                // 直接保存文件
                await handleFileSave(result.workbook, result.themeName, ugcResult, allObstacleResult);
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
            console.log(`当前映射模式: ${currentMappingMode}`);

            let updateResult;
            if (currentMappingMode === 'direct') {
                console.log('使用直接映射模式处理颜色数据...');
                updateResult = updateThemeColorsDirect(themeRowIndex, themeName, isNewTheme);
            } else {
                console.log('使用JSON映射模式处理颜色数据...');
                updateResult = updateThemeColors(themeRowIndex, themeName);
            }

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

        // 🔧 修复：找到最后一个有效数据行，避免在空行后添加
        let lastValidRowIndex = data.length - 1;
        while (lastValidRowIndex > 0 && (!data[lastValidRowIndex] || data[lastValidRowIndex].every(cell => !cell || cell === ''))) {
            lastValidRowIndex--;
        }

        const newRowIndex = lastValidRowIndex + 1;
        const newRow = new Array(headerRow.length).fill('');

        console.log(`最后有效行索引: ${lastValidRowIndex}`);
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

        // 🔧 修复：如果新行索引小于当前数据长度，则替换现有空行；否则添加新行
        if (newRowIndex < data.length) {
            data[newRowIndex] = newRow;
            console.log(`✅ 新行已替换空行，索引: ${newRowIndex}`);
        } else {
            data.push(newRow);
            console.log(`✅ 新行已添加到数据数组，当前总行数: ${data.length}`);
        }

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
     * 从原RSC_Theme文件的Color工作表中读取颜色值
     * @param {string} colorChannel - 颜色通道名称（如P1, G1等）
     * @param {string} themeName - 主题名称
     * @returns {string|null} 颜色值或null
     */
    function findColorValueFromRSCTheme(colorChannel, themeName) {
        console.log(`=== 从RSC_Theme文件查找颜色值: ${colorChannel} ===`);

        if (!rscThemeData || !rscThemeData.workbook) {
            console.warn('RSC_Theme数据不可用');
            return null;
        }

        try {
            // 检查是否有Color工作表
            const colorSheetName = 'Color';
            if (!rscThemeData.workbook.SheetNames.includes(colorSheetName)) {
                console.warn('RSC_Theme文件中未找到Color工作表');
                return null;
            }

            const colorSheet = rscThemeData.workbook.Sheets[colorSheetName];
            const colorData = XLSX.utils.sheet_to_json(colorSheet, {
                header: 1,
                defval: '',
                raw: false
            });

            if (!colorData || colorData.length < 2) {
                console.warn('RSC_Theme Color工作表数据不足');
                return null;
            }

            const headers = colorData[0];
            console.log('RSC_Theme Color工作表表头:', headers);

            // 查找颜色通道列索引
            const targetChannel = colorChannel.toUpperCase();
            const channelColumnIndex = headers.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim().toUpperCase();
                return headerStr === targetChannel;
            });

            if (channelColumnIndex === -1) {
                console.log(`RSC_Theme中未找到颜色通道 ${colorChannel}`);
                return null;
            }

            // 查找notes列索引
            const notesColumnIndex = headers.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim().toLowerCase();
                return headerStr === 'notes';
            });

            if (notesColumnIndex === -1) {
                console.warn('RSC_Theme Color工作表中未找到notes列');
                return null;
            }

            // 查找对应主题的行
            for (let i = 1; i < colorData.length; i++) {
                const row = colorData[i];
                const rowThemeName = row[notesColumnIndex];

                if (rowThemeName === themeName) {
                    const colorValue = row[channelColumnIndex];

                    if (colorValue && colorValue !== '' && colorValue !== null && colorValue !== undefined) {
                        // 清理颜色值
                        let cleanValue = colorValue.toString().trim().toUpperCase();
                        if (cleanValue.startsWith('#')) {
                            cleanValue = cleanValue.substring(1);
                        }

                        console.log(`✅ 从RSC_Theme找到 ${colorChannel} = ${cleanValue} (主题: ${themeName})`);
                        return cleanValue;
                    }
                }
            }

            console.log(`RSC_Theme中未找到主题 ${themeName} 的 ${colorChannel} 颜色值`);
            return null;

        } catch (error) {
            console.error(`从RSC_Theme查找 ${colorChannel} 时出错:`, error);
            return null;
        }
    }

    /**
     * 直接映射模式：在源数据中查找颜色值
     * @param {string} colorChannel - 颜色通道名称（如P1, G1等）
     * @param {boolean} isNewTheme - 是否为新建主题
     * @param {string} themeName - 主题名称
     * @returns {string|null} 颜色值或null
     */
    function findColorValueDirect(colorChannel, isNewTheme = false, themeName = '') {
        console.log(`=== 直接映射查找颜色值: ${colorChannel} (新主题: ${isNewTheme}, 主题名: ${themeName}) ===`);

        if (!sourceData || !sourceData.workbook) {
            console.warn('源数据不可用');
            return null;
        }

        // 解析Status工作表状态
        const statusInfo = parseStatusSheet(sourceData);
        console.log('Status状态信息:', statusInfo);

        if (!statusInfo.hasColorField) {
            console.warn('Status工作表中没有Color字段，根据主题类型处理');

            if (!isNewTheme && themeName) {
                // 更新现有主题模式：直接从RSC_Theme文件读取
                console.log('更新现有主题且无Color字段，直接从RSC_Theme文件读取颜色值');
                const rscColorValue = findColorValueFromRSCTheme(colorChannel, themeName);
                if (rscColorValue) {
                    console.log(`✅ 从RSC_Theme文件找到: ${colorChannel} = ${rscColorValue}`);
                    return rscColorValue;
                }
            }

            // 新建主题模式或未找到：返回null，使用默认值
            console.log(`⚠️ 无Color字段，${isNewTheme ? '新建主题' : '现有主题'}未找到颜色值: ${colorChannel}`);
            return null;
        }

        const isColorValid = statusInfo.isColorValid;
        console.log(`Color字段状态: ${isColorValid ? '有效(1)' : '无效(0)'}`);

        // 根据Color状态和主题类型决定处理逻辑
        if (isColorValid) {
            // Color状态为有效(1)
            console.log('Color状态有效，优先从源数据Color工作表查找');

            // 优先从源数据Color工作表查找
            const sourceColorValue = findColorValueFromSourceColor(colorChannel);
            if (sourceColorValue) {
                console.log(`✅ 从源数据Color工作表找到: ${colorChannel} = ${sourceColorValue}`);
                return sourceColorValue;
            }

            if (!isNewTheme && themeName) {
                // 更新现有主题模式：回退到RSC_Theme文件
                console.log('源数据Color工作表未找到，回退到RSC_Theme文件查找');
                const rscColorValue = findColorValueFromRSCTheme(colorChannel, themeName);
                if (rscColorValue) {
                    console.log(`✅ 从RSC_Theme文件找到: ${colorChannel} = ${rscColorValue}`);
                    return rscColorValue;
                }
            }

            console.log(`⚠️ Color状态有效但未找到颜色值: ${colorChannel}`);
            return null;

        } else {
            // Color状态为无效(0)
            console.log('Color状态无效，忽略源数据Color工作表');

            if (!isNewTheme && themeName) {
                // 更新现有主题模式：直接从RSC_Theme文件读取
                console.log('直接从RSC_Theme文件读取颜色值');
                const rscColorValue = findColorValueFromRSCTheme(colorChannel, themeName);
                if (rscColorValue) {
                    console.log(`✅ 从RSC_Theme文件找到: ${colorChannel} = ${rscColorValue}`);
                    return rscColorValue;
                }
            }

            // 新建主题模式或未找到：返回null，使用默认值
            console.log(`⚠️ Color状态无效，${isNewTheme ? '新建主题' : '现有主题'}未找到颜色值: ${colorChannel}`);
            return null;
        }
    }

    /**
     * 从源数据Color工作表中查找颜色值
     * @param {string} colorChannel - 颜色通道名称（如P1, G1等）
     * @returns {string|null} 颜色值或null
     */
    function findColorValueFromSourceColor(colorChannel) {
        try {
            // 读取Color工作表
            const colorSheet = sourceData.workbook.Sheets['Color'];
            if (!colorSheet) {
                console.log('源数据中Color工作表不存在');
                return null;
            }

            const colorData = XLSX.utils.sheet_to_json(colorSheet, {
                header: 1,
                defval: '',
                raw: false
            });

            if (!colorData || colorData.length < 2) {
                console.log('源数据Color工作表数据不足');
                return null;
            }

            const headers = colorData[0];
            const dataRow = colorData[1]; // 第二行是数据行

            // 查找对应的列索引 - 灵活匹配
            const targetChannel = colorChannel.toUpperCase();
            const columnIndex = headers.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim().toUpperCase();
                return headerStr === targetChannel;
            });

            if (columnIndex === -1) {
                console.log(`源数据Color工作表：未找到字段 ${colorChannel}，可用字段: ${headers.join(', ')}`);
                return null;
            }

            const colorValue = dataRow[columnIndex];

            // 简化验证：只检查非空
            if (!colorValue || colorValue === '' || colorValue === null || colorValue === undefined) {
                console.log(`源数据Color工作表：字段 ${colorChannel} 的值为空，原始值: "${colorValue}"`);
                return null;
            }

            // 清理颜色值（移除#号，转换为大写）
            let cleanValue = colorValue.toString().trim().toUpperCase();
            if (cleanValue.startsWith('#')) {
                cleanValue = cleanValue.substring(1);
            }

            console.log(`源数据Color工作表：找到 ${colorChannel} = ${cleanValue} (列索引: ${columnIndex})`);
            return cleanValue;

        } catch (error) {
            console.error(`从源数据Color工作表查找 ${colorChannel} 时出错:`, error);
            return null;
        }
    }

    /**
     * 从源数据Light工作表中查找Light字段值
     * @param {string} lightField - Light字段名称（如Max, Dark, Min等）
     * @returns {string|null} Light字段值或null
     */
    function findLightValueFromSourceLight(lightField) {
        try {
            // 读取Light工作表
            const lightSheet = sourceData.workbook.Sheets['Light'];
            if (!lightSheet) {
                console.log('源数据中Light工作表不存在');
                return null;
            }

            const lightData = XLSX.utils.sheet_to_json(lightSheet, {
                header: 1,
                defval: '',
                raw: false
            });

            if (!lightData || lightData.length < 2) {
                console.log('源数据Light工作表数据不足');
                return null;
            }

            const headers = lightData[0];
            const dataRow = lightData[1]; // 第二行是数据行

            // 查找对应的列索引 - 灵活匹配
            const targetField = lightField.toString().trim();
            const columnIndex = headers.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim();
                return headerStr === targetField;
            });

            if (columnIndex === -1) {
                console.log(`源数据Light工作表：未找到字段 ${lightField}，可用字段: ${headers.join(', ')}`);
                return null;
            }

            const lightValue = dataRow[columnIndex];

            // 验证：检查非空
            if (lightValue === '' || lightValue === null || lightValue === undefined) {
                console.log(`源数据Light工作表：字段 ${lightField} 的值为空，原始值: "${lightValue}"`);
                return null;
            }

            // 清理Light值（转换为字符串）
            let cleanValue = lightValue.toString().trim();

            console.log(`✅ 从源数据Light工作表找到: ${lightField} = ${cleanValue}`);
            return cleanValue;

        } catch (error) {
            console.error(`从源数据Light工作表查找 ${lightField} 时出错:`, error);
            return null;
        }
    }

    /**
     * 从源数据ColorInfo工作表中查找ColorInfo字段值
     * @param {string} colorInfoField - ColorInfo字段名称
     * @returns {string|null} ColorInfo字段值，未找到返回null
     */
    function findColorInfoValueFromSourceColorInfo(colorInfoField) {
        console.log(`=== 开始从源数据ColorInfo工作表查找字段: ${colorInfoField} ===`);

        if (!sourceData || !sourceData.workbook) {
            console.warn('源数据未加载，无法从ColorInfo工作表读取');
            return null;
        }

        try {
            const workbook = sourceData.workbook;
            const sheetNames = workbook.SheetNames;
            console.log('源数据包含的工作表:', sheetNames);

            // 查找ColorInfo工作表
            if (!sheetNames.includes('ColorInfo')) {
                console.log('源数据中未找到ColorInfo工作表');
                return null;
            }

            const colorInfoWorksheet = workbook.Sheets['ColorInfo'];
            const colorInfoData = XLSX.utils.sheet_to_json(colorInfoWorksheet, {
                header: 1,
                defval: '',
                raw: false
            });

            if (colorInfoData.length === 0) {
                console.log('ColorInfo工作表为空');
                return null;
            }

            const headerRow = colorInfoData[0];
            console.log('ColorInfo工作表表头:', headerRow);

            // 查找目标字段列
            const fieldColumnIndex = headerRow.findIndex(col => col === colorInfoField);
            if (fieldColumnIndex === -1) {
                console.log(`ColorInfo工作表中未找到字段: ${colorInfoField}`);
                return null;
            }

            // 获取字段值（通常在第二行）
            if (colorInfoData.length > 1) {
                const fieldValue = colorInfoData[1][fieldColumnIndex];
                if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                    console.log(`✅ 从源数据ColorInfo工作表找到: ${colorInfoField} = ${fieldValue}`);
                    return fieldValue.toString();
                }
            }

            console.log(`ColorInfo工作表中字段 ${colorInfoField} 值为空`);
            return null;

        } catch (error) {
            console.error('从源数据ColorInfo工作表读取字段时出错:', error);
            return null;
        }
    }

    /**
     * 从RSC_Theme ColorInfo工作表中查找ColorInfo字段值
     * @param {string} colorInfoField - ColorInfo字段名称
     * @param {string} themeName - 主题名称
     * @returns {string|null} ColorInfo字段值，未找到返回null
     */
    function findColorInfoValueFromRSCThemeColorInfo(colorInfoField, themeName) {
        console.log(`=== 开始从RSC_Theme ColorInfo工作表查找字段: ${colorInfoField}, 主题: ${themeName} ===`);

        if (!rscAllSheetsData || !rscAllSheetsData['ColorInfo']) {
            console.warn('RSC_Theme ColorInfo数据未加载');
            return null;
        }

        try {
            const colorInfoData = rscAllSheetsData['ColorInfo'];
            if (colorInfoData.length === 0) {
                console.log('RSC_Theme ColorInfo工作表为空');
                return null;
            }

            const headerRow = colorInfoData[0];
            console.log('RSC_Theme ColorInfo工作表表头:', headerRow);

            // 查找notes列
            const notesColumnIndex = headerRow.findIndex(col => col === 'notes');
            if (notesColumnIndex === -1) {
                console.log('RSC_Theme ColorInfo工作表中未找到notes列');
                return null;
            }

            // 查找目标字段列
            const fieldColumnIndex = headerRow.findIndex(col => col === colorInfoField);
            if (fieldColumnIndex === -1) {
                console.log(`RSC_Theme ColorInfo工作表中未找到字段: ${colorInfoField}`);
                return null;
            }

            // 查找主题对应的行
            for (let i = 1; i < colorInfoData.length; i++) {
                const row = colorInfoData[i];
                if (row[notesColumnIndex] === themeName) {
                    const fieldValue = row[fieldColumnIndex];
                    if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                        console.log(`✅ 从RSC_Theme ColorInfo工作表找到: ${colorInfoField} = ${fieldValue} (主题: ${themeName})`);
                        return fieldValue.toString();
                    } else {
                        console.log(`RSC_Theme ColorInfo工作表中主题 ${themeName} 的字段 ${colorInfoField} 值为空`);
                        return null;
                    }
                }
            }

            console.log(`RSC_Theme ColorInfo工作表中未找到主题: ${themeName}`);
            return null;

        } catch (error) {
            console.error('从RSC_Theme ColorInfo工作表读取字段时出错:', error);
            return null;
        }
    }

    /**
     * 从RSC_Theme ColorInfo工作表读取第一个主题的字段值（用于新建主题）
     * @param {string} colorInfoField - ColorInfo字段名称
     * @returns {string|null} ColorInfo字段值或null
     */
    function findColorInfoValueFromRSCThemeColorInfoFirstTheme(colorInfoField) {
        console.log(`=== 从RSC_Theme ColorInfo工作表读取第一个主题的字段值: ${colorInfoField} ===`);

        if (!rscAllSheetsData || !rscAllSheetsData['ColorInfo']) {
            console.warn('RSC_Theme ColorInfo数据未加载');
            return null;
        }

        try {
            const colorInfoData = rscAllSheetsData['ColorInfo'];

            // 检查数据是否足够（需要至少6行：表头+5行元数据+第一个主题）
            if (colorInfoData.length <= 5) {
                console.warn('RSC_Theme ColorInfo工作表数据不足（需要至少6行）');
                return null;
            }

            const headerRow = colorInfoData[0];
            console.log('RSC_Theme ColorInfo工作表表头:', headerRow);

            // 查找目标字段列
            const fieldColumnIndex = headerRow.findIndex(col => col === colorInfoField);
            if (fieldColumnIndex === -1) {
                console.log(`RSC_Theme ColorInfo工作表中未找到字段: ${colorInfoField}`);
                return null;
            }

            // 读取第一个主题（行索引5，第6行）
            const firstThemeRowIndex = 5;
            const firstThemeRow = colorInfoData[firstThemeRowIndex];

            if (!firstThemeRow) {
                console.warn(`RSC_Theme ColorInfo工作表第一个主题行不存在（行索引: ${firstThemeRowIndex}）`);
                return null;
            }

            const fieldValue = firstThemeRow[fieldColumnIndex];

            if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                console.log(`✅ 从RSC_Theme ColorInfo工作表第一个主题找到: ${colorInfoField} = ${fieldValue}`);
                return fieldValue.toString();
            } else {
                console.log(`RSC_Theme ColorInfo工作表第一个主题中未找到 ${colorInfoField} 字段值`);
                return null;
            }

        } catch (error) {
            console.error('从RSC_Theme ColorInfo工作表读取第一个主题的字段时出错:', error);
            return null;
        }
    }

    /**
     * 从RSC_Theme Light工作表中读取Light字段值
     * @param {string} lightField - Light字段名称（如Max, Dark, Min等）
     * @param {string} themeName - 主题名称
     * @returns {string|null} Light字段值或null
     */
    function findLightValueFromRSCThemeLight(lightField, themeName) {
        console.log(`=== 从RSC_Theme Light工作表查找Light字段值: ${lightField} ===`);

        if (!rscAllSheetsData || !rscAllSheetsData['Light']) {
            console.warn('RSC_Theme Light数据不可用');
            return null;
        }

        try {
            const lightData = rscAllSheetsData['Light'];
            const lightHeaderRow = lightData[0];

            if (!lightData || lightData.length < 2) {
                console.warn('RSC_Theme Light工作表数据不足');
                return null;
            }

            console.log('RSC_Theme Light工作表表头:', lightHeaderRow);

            // 查找Light字段列索引
            const targetField = lightField.toString().trim();
            const fieldColumnIndex = lightHeaderRow.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim();
                return headerStr === targetField;
            });

            if (fieldColumnIndex === -1) {
                console.log(`RSC_Theme Light工作表中未找到字段 ${lightField}`);
                return null;
            }

            // 查找notes列索引
            const notesColumnIndex = lightHeaderRow.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim().toLowerCase();
                return headerStr === 'notes';
            });

            if (notesColumnIndex === -1) {
                console.warn('RSC_Theme Light工作表中未找到notes列');
                return null;
            }

            // 查找对应主题的行
            for (let i = 1; i < lightData.length; i++) {
                const row = lightData[i];
                const rowThemeName = row[notesColumnIndex];

                if (rowThemeName === themeName) {
                    const lightValue = row[fieldColumnIndex];

                    if (lightValue !== undefined && lightValue !== null && lightValue !== '') {
                        // 清理Light值
                        let cleanValue = lightValue.toString().trim();

                        console.log(`✅ 从RSC_Theme Light工作表找到 ${lightField} = ${cleanValue} (主题: ${themeName})`);
                        return cleanValue;
                    }
                }
            }

            console.log(`RSC_Theme Light工作表中未找到主题 ${themeName} 的 ${lightField} 字段值`);
            return null;

        } catch (error) {
            console.error(`从RSC_Theme Light工作表查找 ${lightField} 时出错:`, error);
            return null;
        }
    }

    /**
     * 从RSC_Theme Light工作表读取第一个主题的字段值（用于新建主题）
     * @param {string} lightField - Light字段名称
     * @returns {string|null} Light字段值或null
     */
    function findLightValueFromRSCThemeLightFirstTheme(lightField) {
        console.log(`=== 从RSC_Theme Light工作表读取第一个主题的字段值: ${lightField} ===`);

        if (!rscAllSheetsData || !rscAllSheetsData['Light']) {
            console.warn('RSC_Theme Light数据不可用');
            return null;
        }

        try {
            const lightData = rscAllSheetsData['Light'];
            const lightHeaderRow = lightData[0];

            // 检查数据是否足够（需要至少6行：表头+5行元数据+第一个主题）
            if (!lightData || lightData.length <= 5) {
                console.warn('RSC_Theme Light工作表数据不足（需要至少6行）');
                return null;
            }

            console.log('RSC_Theme Light工作表表头:', lightHeaderRow);

            // 查找Light字段列索引
            const targetField = lightField.toString().trim();
            const fieldColumnIndex = lightHeaderRow.findIndex(header => {
                if (!header) return false;
                const headerStr = header.toString().trim();
                return headerStr === targetField;
            });

            if (fieldColumnIndex === -1) {
                console.log(`RSC_Theme Light工作表中未找到字段 ${lightField}`);
                return null;
            }

            // 读取第一个主题（行索引5，第6行）
            const firstThemeRowIndex = 5;
            const firstThemeRow = lightData[firstThemeRowIndex];

            if (!firstThemeRow) {
                console.warn(`RSC_Theme Light工作表第一个主题行不存在（行索引: ${firstThemeRowIndex}）`);
                return null;
            }

            const lightValue = firstThemeRow[fieldColumnIndex];

            if (lightValue !== undefined && lightValue !== null && lightValue !== '') {
                // 清理Light值
                let cleanValue = lightValue.toString().trim();

                console.log(`✅ 从RSC_Theme Light工作表第一个主题找到 ${lightField} = ${cleanValue}`);
                return cleanValue;
            }

            console.log(`RSC_Theme Light工作表第一个主题中未找到 ${lightField} 字段值`);
            return null;

        } catch (error) {
            console.error(`从RSC_Theme Light工作表读取第一个主题的 ${lightField} 时出错:`, error);
            return null;
        }
    }

    /**
     * 在直接映射模式下查找ColorInfo字段值（带条件判断）
     * @param {string} colorInfoField - ColorInfo字段名称
     * @param {boolean} isNewTheme - 是否为新建主题
     * @param {string} themeName - 主题名称
     * @returns {string|null} ColorInfo字段值或null
     */
    function findColorInfoValueDirect(colorInfoField, isNewTheme = false, themeName = '') {
        console.log(`=== 直接映射模式查找ColorInfo字段: ${colorInfoField} ===`);
        console.log(`主题类型: ${isNewTheme ? '新建主题' : '更新现有主题'}, 主题名称: ${themeName}`);

        if (!sourceData || !sourceData.workbook) {
            console.warn('源数据未加载，无法进行ColorInfo字段条件读取');
            return null;
        }

        // ✅ 特殊规则：钻石颜色字段始终读取第一个主题（无论Status状态如何）
        const diamondColorFields = ['PickupDiffR', 'PickupDiffG', 'PickupDiffB', 'PickupReflR', 'PickupReflG', 'PickupReflB', 'BallSpecR', 'BallSpecG', 'BallSpecB'];
        if (diamondColorFields.includes(colorInfoField)) {
            console.log(`✅ 钻石颜色字段 ${colorInfoField}：特殊规则，始终读取第一个主题`);
            if (isNewTheme) {
                const rscColorInfoValue = findColorInfoValueFromRSCThemeColorInfoFirstTheme(colorInfoField);
                if (rscColorInfoValue) {
                    console.log(`✅ 从RSC_Theme ColorInfo工作表第一个主题找到钻石颜色: ${colorInfoField} = ${rscColorInfoValue}`);
                    return rscColorInfoValue;
                }
            } else if (themeName) {
                const rscColorInfoValue = findColorInfoValueFromRSCThemeColorInfo(colorInfoField, themeName);
                if (rscColorInfoValue) {
                    console.log(`✅ 从RSC_Theme ColorInfo工作表找到钻石颜色: ${colorInfoField} = ${rscColorInfoValue}`);
                    return rscColorInfoValue;
                }
            }
            return null;
        }

        // 解析Status工作表获取ColorInfo状态
        const statusInfo = parseStatusSheet(sourceData);
        console.log('ColorInfo状态信息:', {
            hasColorInfoField: statusInfo.hasColorInfoField,
            colorInfoStatus: statusInfo.colorInfoStatus,
            isColorInfoValid: statusInfo.isColorInfoValid
        });

        if (!statusInfo.hasColorInfoField) {
            console.warn('Status工作表中没有ColorInfo字段，根据主题类型处理');

            if (isNewTheme) {
                // ✅ 新建主题模式：从RSC_Theme ColorInfo工作表读取第一个主题的数据
                console.log('新建主题且无ColorInfo字段，从RSC_Theme ColorInfo工作表读取第一个主题的字段值');
                const rscColorInfoValue = findColorInfoValueFromRSCThemeColorInfoFirstTheme(colorInfoField);
                if (rscColorInfoValue) {
                    console.log(`✅ 从RSC_Theme ColorInfo工作表第一个主题找到: ${colorInfoField} = ${rscColorInfoValue}`);
                    return rscColorInfoValue;
                }
            } else if (themeName) {
                // 更新现有主题模式：直接从RSC_Theme ColorInfo工作表读取
                console.log('更新现有主题且无ColorInfo字段，直接从RSC_Theme ColorInfo工作表读取字段值');
                const rscColorInfoValue = findColorInfoValueFromRSCThemeColorInfo(colorInfoField, themeName);
                if (rscColorInfoValue) {
                    console.log(`✅ 从RSC_Theme ColorInfo工作表找到: ${colorInfoField} = ${rscColorInfoValue}`);
                    return rscColorInfoValue;
                }
            }

            console.log('无ColorInfo字段且无法从RSC_Theme读取，返回null');
            return null;
        }

        if (statusInfo.isColorInfoValid) {
            // ColorInfo状态为有效(1)
            console.log('ColorInfo状态有效，尝试从源数据ColorInfo工作表读取');

            const sourceColorInfoValue = findColorInfoValueFromSourceColorInfo(colorInfoField);
            if (sourceColorInfoValue) {
                console.log(`✅ 从源数据ColorInfo工作表找到: ${colorInfoField} = ${sourceColorInfoValue}`);
                return sourceColorInfoValue;
            }

            // 如果从源数据ColorInfo工作表读取字段值没有找到字段，回退到RSC_Theme ColorInfo工作表
            if (isNewTheme) {
                // ✅ 新建主题模式：回退到RSC_Theme ColorInfo工作表的第一个主题
                console.log('新建主题且源数据ColorInfo工作表未找到字段，回退到RSC_Theme ColorInfo工作表的第一个主题查找');
                const rscColorInfoValue = findColorInfoValueFromRSCThemeColorInfoFirstTheme(colorInfoField);
                if (rscColorInfoValue) {
                    console.log(`✅ 从RSC_Theme ColorInfo工作表第一个主题找到: ${colorInfoField} = ${rscColorInfoValue}`);
                    return rscColorInfoValue;
                }
            } else if (themeName) {
                console.log('源数据ColorInfo工作表未找到字段，回退到RSC_Theme ColorInfo工作表查找');
                const rscColorInfoValue = findColorInfoValueFromRSCThemeColorInfo(colorInfoField, themeName);
                if (rscColorInfoValue) {
                    console.log(`✅ 从RSC_Theme ColorInfo工作表找到: ${colorInfoField} = ${rscColorInfoValue}`);
                    return rscColorInfoValue;
                }
            }

            console.log('ColorInfo状态有效但未找到字段值，返回null');
            return null;
        } else {
            // ColorInfo状态为无效(0)
            console.log('ColorInfo状态无效，忽略源数据ColorInfo工作表');

            if (isNewTheme) {
                // ✅ 新建主题模式：从RSC_Theme ColorInfo工作表读取第一个主题的数据
                console.log('新建主题且ColorInfo状态无效，从RSC_Theme ColorInfo工作表读取第一个主题的字段值');
                const rscColorInfoValue = findColorInfoValueFromRSCThemeColorInfoFirstTheme(colorInfoField);
                if (rscColorInfoValue) {
                    console.log(`✅ 从RSC_Theme ColorInfo工作表第一个主题找到: ${colorInfoField} = ${rscColorInfoValue}`);
                    return rscColorInfoValue;
                }
            } else if (themeName) {
                // 更新现有主题模式：直接从RSC_Theme ColorInfo工作表读取
                console.log('直接从RSC_Theme ColorInfo工作表读取ColorInfo字段值');
                const rscColorInfoValue = findColorInfoValueFromRSCThemeColorInfo(colorInfoField, themeName);
                if (rscColorInfoValue) {
                    console.log(`✅ 从RSC_Theme ColorInfo工作表找到: ${colorInfoField} = ${rscColorInfoValue}`);
                    return rscColorInfoValue;
                }
            }

            console.log('ColorInfo状态无效且无法从RSC_Theme读取，返回null');
            return null;
        }
    }

    /**
     * 从源数据VolumetricFog工作表中读取VolumetricFog字段值
     * @param {string} volumetricFogField - VolumetricFog字段名称（如Color, X, Y等）
     * @returns {string|null} VolumetricFog字段值或null
     */
    function findVolumetricFogValueFromSourceVolumetricFog(volumetricFogField) {
        console.log(`=== 从源数据VolumetricFog工作表查找字段: ${volumetricFogField} ===`);

        try {
            if (!sourceData || !sourceData.workbook) {
                console.warn('源数据未加载，无法从源数据VolumetricFog工作表读取字段');
                return null;
            }

            const volumetricFogSheetName = 'VolumetricFog';
            const volumetricFogWorksheet = sourceData.workbook.Sheets[volumetricFogSheetName];

            if (!volumetricFogWorksheet) {
                console.log(`源数据中未找到${volumetricFogSheetName}工作表`);
                return null;
            }

            // 将工作表转换为数组
            const volumetricFogData = XLSX.utils.sheet_to_json(volumetricFogWorksheet, { header: 1 });

            if (!volumetricFogData || volumetricFogData.length < 2) {
                console.log(`${volumetricFogSheetName}工作表数据不足`);
                return null;
            }

            // 查找字段列索引
            const headerRow = volumetricFogData[0];
            const fieldColumnIndex = headerRow.findIndex(col => col === volumetricFogField);

            if (fieldColumnIndex === -1) {
                console.log(`源数据${volumetricFogSheetName}工作表中未找到字段: ${volumetricFogField}`);
                return null;
            }

            // 从第二行获取字段值（假设只有一行数据）
            if (volumetricFogData.length > 1) {
                const fieldValue = volumetricFogData[1][fieldColumnIndex];
                if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                    console.log(`✅ 从源数据${volumetricFogSheetName}工作表找到: ${volumetricFogField} = ${fieldValue}`);
                    return fieldValue.toString();
                } else {
                    console.log(`源数据${volumetricFogSheetName}工作表中字段 ${volumetricFogField} 值为空`);
                    return null;
                }
            }

            console.log(`源数据${volumetricFogSheetName}工作表中没有数据行`);
            return null;

        } catch (error) {
            console.error('从源数据VolumetricFog工作表读取字段时出错:', error);
            return null;
        }
    }

    /**
     * 从RSC_Theme VolumetricFog工作表中读取VolumetricFog字段值
     * @param {string} volumetricFogField - VolumetricFog字段名称（如Color, X, Y等）
     * @param {string} themeName - 主题名称
     * @returns {string|null} VolumetricFog字段值或null
     */
    function findVolumetricFogValueFromRSCThemeVolumetricFog(volumetricFogField, themeName) {
        console.log(`=== 从RSC_Theme VolumetricFog工作表查找字段: ${volumetricFogField}, 主题: ${themeName} ===`);

        try {
            if (!rscAllSheetsData || !rscAllSheetsData['VolumetricFog']) {
                console.warn('RSC_Theme VolumetricFog数据未加载');
                return null;
            }

            const volumetricFogData = rscAllSheetsData['VolumetricFog'];

            if (!volumetricFogData || volumetricFogData.length < 2) {
                console.log('RSC_Theme VolumetricFog工作表数据不足');
                return null;
            }

            // 查找字段列索引和notes列索引
            const headerRow = volumetricFogData[0];
            const fieldColumnIndex = headerRow.findIndex(col => col === volumetricFogField);
            const notesColumnIndex = headerRow.findIndex(col => col === 'notes');

            if (fieldColumnIndex === -1) {
                console.log(`RSC_Theme VolumetricFog工作表中未找到字段: ${volumetricFogField}`);
                return null;
            }

            if (notesColumnIndex === -1) {
                console.log('RSC_Theme VolumetricFog工作表中未找到notes列');
                return null;
            }

            // 查找主题对应的行
            for (let i = 1; i < volumetricFogData.length; i++) {
                const row = volumetricFogData[i];
                if (row[notesColumnIndex] === themeName) {
                    const fieldValue = row[fieldColumnIndex];
                    if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                        console.log(`✅ 从RSC_Theme VolumetricFog工作表找到: ${volumetricFogField} = ${fieldValue} (主题: ${themeName})`);
                        return fieldValue.toString();
                    } else {
                        console.log(`RSC_Theme VolumetricFog工作表中主题 ${themeName} 的字段 ${volumetricFogField} 值为空`);
                        return null;
                    }
                }
            }

            console.log(`RSC_Theme VolumetricFog工作表中未找到主题: ${themeName}`);
            return null;

        } catch (error) {
            console.error('从RSC_Theme VolumetricFog工作表读取字段时出错:', error);
            return null;
        }
    }

    /**
     * 从RSC_Theme VolumetricFog工作表读取第一个主题的字段值（用于新建主题）
     * @param {string} volumetricFogField - VolumetricFog字段名称
     * @returns {string|null} VolumetricFog字段值或null
     */
    function findVolumetricFogValueFromRSCThemeVolumetricFogFirstTheme(volumetricFogField) {
        console.log(`=== 从RSC_Theme VolumetricFog工作表读取第一个主题的字段值: ${volumetricFogField} ===`);

        try {
            if (!rscAllSheetsData || !rscAllSheetsData['VolumetricFog']) {
                console.warn('RSC_Theme VolumetricFog数据未加载');
                return null;
            }

            const volumetricFogData = rscAllSheetsData['VolumetricFog'];

            // 检查数据是否足够（需要至少6行：表头+5行元数据+第一个主题）
            if (!volumetricFogData || volumetricFogData.length <= 5) {
                console.warn('RSC_Theme VolumetricFog工作表数据不足（需要至少6行）');
                return null;
            }

            // 查找字段列索引
            const headerRow = volumetricFogData[0];
            const fieldColumnIndex = headerRow.findIndex(col => col === volumetricFogField);

            if (fieldColumnIndex === -1) {
                console.log(`RSC_Theme VolumetricFog工作表中未找到字段: ${volumetricFogField}`);
                return null;
            }

            // 读取第一个主题（行索引5，第6行）
            const firstThemeRowIndex = 5;
            const firstThemeRow = volumetricFogData[firstThemeRowIndex];

            if (!firstThemeRow) {
                console.warn(`RSC_Theme VolumetricFog工作表第一个主题行不存在（行索引: ${firstThemeRowIndex}）`);
                return null;
            }

            const fieldValue = firstThemeRow[fieldColumnIndex];

            if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                console.log(`✅ 从RSC_Theme VolumetricFog工作表第一个主题找到: ${volumetricFogField} = ${fieldValue}`);
                return fieldValue.toString();
            } else {
                console.log(`RSC_Theme VolumetricFog工作表第一个主题中未找到 ${volumetricFogField} 字段值`);
                return null;
            }

        } catch (error) {
            console.error('从RSC_Theme VolumetricFog工作表读取第一个主题的字段时出错:', error);
            return null;
        }
    }

    /**
     * 直接映射模式：VolumetricFog字段条件读取
     * @param {string} volumetricFogField - VolumetricFog字段名称（如Color, X, Y等）
     * @param {boolean} isNewTheme - 是否为新建主题
     * @param {string} themeName - 主题名称
     * @returns {string|null} VolumetricFog字段值或null
     */
    function findVolumetricFogValueDirect(volumetricFogField, isNewTheme = false, themeName = '') {
        console.log(`=== 直接映射模式查找VolumetricFog字段: ${volumetricFogField} ===`);
        console.log(`主题类型: ${isNewTheme ? '新建主题' : '更新现有主题'}, 主题名称: ${themeName}`);

        if (!sourceData || !sourceData.workbook) {
            console.warn('源数据未加载，无法进行VolumetricFog字段条件读取');
            return null;
        }

        // 解析Status工作表获取VolumetricFog状态
        const statusInfo = parseStatusSheet(sourceData);
        console.log('VolumetricFog状态信息:', {
            hasVolumetricFogField: statusInfo.hasVolumetricFogField,
            volumetricFogStatus: statusInfo.volumetricFogStatus,
            isVolumetricFogValid: statusInfo.isVolumetricFogValid
        });

        if (!statusInfo.hasVolumetricFogField) {
            console.warn('Status工作表中没有VolumetricFog字段，根据主题类型处理');

            if (isNewTheme) {
                // ✅ 新建主题模式：从RSC_Theme VolumetricFog工作表读取第一个主题的数据
                console.log('新建主题且无VolumetricFog字段，从RSC_Theme VolumetricFog工作表读取第一个主题的字段值');
                const rscVolumetricFogValue = findVolumetricFogValueFromRSCThemeVolumetricFogFirstTheme(volumetricFogField);
                if (rscVolumetricFogValue) {
                    console.log(`✅ 从RSC_Theme VolumetricFog工作表第一个主题找到: ${volumetricFogField} = ${rscVolumetricFogValue}`);
                    return rscVolumetricFogValue;
                }
            } else if (themeName) {
                // 更新现有主题模式：直接从RSC_Theme VolumetricFog工作表读取
                console.log('更新现有主题且无VolumetricFog字段，直接从RSC_Theme VolumetricFog工作表读取字段值');
                const rscVolumetricFogValue = findVolumetricFogValueFromRSCThemeVolumetricFog(volumetricFogField, themeName);
                if (rscVolumetricFogValue) {
                    console.log(`✅ 从RSC_Theme VolumetricFog工作表找到: ${volumetricFogField} = ${rscVolumetricFogValue}`);
                    return rscVolumetricFogValue;
                }
            }

            console.log('无VolumetricFog字段且无法从RSC_Theme读取，返回null');
            return null;
        }

        if (statusInfo.isVolumetricFogValid) {
            // VolumetricFog状态为有效(1)
            console.log('VolumetricFog状态有效，尝试从源数据VolumetricFog工作表读取');

            const sourceVolumetricFogValue = findVolumetricFogValueFromSourceVolumetricFog(volumetricFogField);
            if (sourceVolumetricFogValue) {
                console.log(`✅ 从源数据VolumetricFog工作表找到: ${volumetricFogField} = ${sourceVolumetricFogValue}`);
                return sourceVolumetricFogValue;
            }

            // 如果从源数据VolumetricFog工作表读取字段值没有找到字段，回退到RSC_Theme VolumetricFog工作表
            if (isNewTheme) {
                // ✅ 新建主题模式：回退到RSC_Theme VolumetricFog工作表的第一个主题
                console.log('新建主题且源数据VolumetricFog工作表未找到字段，回退到RSC_Theme VolumetricFog工作表的第一个主题查找');
                const rscVolumetricFogValue = findVolumetricFogValueFromRSCThemeVolumetricFogFirstTheme(volumetricFogField);
                if (rscVolumetricFogValue) {
                    console.log(`✅ 从RSC_Theme VolumetricFog工作表第一个主题找到: ${volumetricFogField} = ${rscVolumetricFogValue}`);
                    return rscVolumetricFogValue;
                }
            } else if (themeName) {
                console.log('源数据VolumetricFog工作表未找到字段，回退到RSC_Theme VolumetricFog工作表查找');
                const rscVolumetricFogValue = findVolumetricFogValueFromRSCThemeVolumetricFog(volumetricFogField, themeName);
                if (rscVolumetricFogValue) {
                    console.log(`✅ 从RSC_Theme VolumetricFog工作表找到: ${volumetricFogField} = ${rscVolumetricFogValue}`);
                    return rscVolumetricFogValue;
                }
            }

            console.log('VolumetricFog状态有效但未找到字段值，返回null');
            return null;
        } else {
            // VolumetricFog状态为无效(0)
            console.log('VolumetricFog状态无效，忽略源数据VolumetricFog工作表');

            if (isNewTheme) {
                // ✅ 新建主题模式：从RSC_Theme VolumetricFog工作表读取第一个主题的数据
                console.log('新建主题且VolumetricFog状态无效，从RSC_Theme VolumetricFog工作表读取第一个主题的字段值');
                const rscVolumetricFogValue = findVolumetricFogValueFromRSCThemeVolumetricFogFirstTheme(volumetricFogField);
                if (rscVolumetricFogValue) {
                    console.log(`✅ 从RSC_Theme VolumetricFog工作表第一个主题找到: ${volumetricFogField} = ${rscVolumetricFogValue}`);
                    return rscVolumetricFogValue;
                }
            } else if (themeName) {
                // 更新现有主题模式：直接从RSC_Theme VolumetricFog工作表读取
                console.log('直接从RSC_Theme VolumetricFog工作表读取VolumetricFog字段值');
                const rscVolumetricFogValue = findVolumetricFogValueFromRSCThemeVolumetricFog(volumetricFogField, themeName);
                if (rscVolumetricFogValue) {
                    console.log(`✅ 从RSC_Theme VolumetricFog工作表找到: ${volumetricFogField} = ${rscVolumetricFogValue}`);
                    return rscVolumetricFogValue;
                }
            }

            console.log('VolumetricFog状态无效且无法从RSC_Theme读取，返回null');
            return null;
        }
    }

    /**
     * 从源数据FloodLight工作表中读取FloodLight字段值
     * @param {string} floodLightField - FloodLight字段名称（如Color, TippingPoint, Strength等）
     * @returns {string|null} FloodLight字段值或null
     */
    function findFloodLightValueFromSourceFloodLight(floodLightField) {
        console.log(`=== 从源数据FloodLight工作表查找字段: ${floodLightField} ===`);

        try {
            if (!sourceData || !sourceData.workbook) {
                console.warn('源数据未加载，无法从源数据FloodLight工作表读取字段');
                return null;
            }

            const floodLightSheetName = 'FloodLight';
            const floodLightWorksheet = sourceData.workbook.Sheets[floodLightSheetName];

            if (!floodLightWorksheet) {
                console.log(`源数据中未找到${floodLightSheetName}工作表`);
                return null;
            }

            // 将工作表转换为数组
            const floodLightData = XLSX.utils.sheet_to_json(floodLightWorksheet, { header: 1 });

            if (!floodLightData || floodLightData.length < 2) {
                console.log(`${floodLightSheetName}工作表数据不足`);
                return null;
            }

            // 查找字段列索引
            const headerRow = floodLightData[0];
            const fieldColumnIndex = headerRow.findIndex(col => col === floodLightField);

            if (fieldColumnIndex === -1) {
                console.log(`源数据${floodLightSheetName}工作表中未找到字段: ${floodLightField}`);
                return null;
            }

            // 从第二行获取字段值（假设只有一行数据）
            if (floodLightData.length > 1) {
                const fieldValue = floodLightData[1][fieldColumnIndex];
                if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                    console.log(`✅ 从源数据${floodLightSheetName}工作表找到: ${floodLightField} = ${fieldValue}`);
                    return fieldValue.toString();
                } else {
                    console.log(`源数据${floodLightSheetName}工作表中字段 ${floodLightField} 值为空`);
                    return null;
                }
            }

            console.log(`源数据${floodLightSheetName}工作表中没有数据行`);
            return null;

        } catch (error) {
            console.error('从源数据FloodLight工作表读取字段时出错:', error);
            return null;
        }
    }

    /**
     * 从RSC_Theme FloodLight工作表中读取FloodLight字段值
     * @param {string} floodLightField - FloodLight字段名称（如Color, TippingPoint, Strength等）
     * @param {string} themeName - 主题名称
     * @returns {string|null} FloodLight字段值或null
     */
    function findFloodLightValueFromRSCThemeFloodLight(floodLightField, themeName) {
        console.log(`=== 从RSC_Theme FloodLight工作表查找字段: ${floodLightField}, 主题: ${themeName} ===`);

        try {
            if (!rscAllSheetsData || !rscAllSheetsData['FloodLight']) {
                console.warn('RSC_Theme FloodLight数据未加载');
                return null;
            }

            const floodLightData = rscAllSheetsData['FloodLight'];

            if (!floodLightData || floodLightData.length < 2) {
                console.log('RSC_Theme FloodLight工作表数据不足');
                return null;
            }

            // 查找字段列索引和notes列索引
            const headerRow = floodLightData[0];
            const fieldColumnIndex = headerRow.findIndex(col => col === floodLightField);
            const notesColumnIndex = headerRow.findIndex(col => col === 'notes');

            if (fieldColumnIndex === -1) {
                console.log(`RSC_Theme FloodLight工作表中未找到字段: ${floodLightField}`);
                return null;
            }

            if (notesColumnIndex === -1) {
                console.log('RSC_Theme FloodLight工作表中未找到notes列');
                return null;
            }

            // 查找主题对应的行
            for (let i = 1; i < floodLightData.length; i++) {
                const row = floodLightData[i];
                if (row[notesColumnIndex] === themeName) {
                    const fieldValue = row[fieldColumnIndex];
                    if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                        console.log(`✅ 从RSC_Theme FloodLight工作表找到: ${floodLightField} = ${fieldValue} (主题: ${themeName})`);
                        return fieldValue.toString();
                    } else {
                        console.log(`RSC_Theme FloodLight工作表中主题 ${themeName} 的字段 ${floodLightField} 值为空`);
                        return null;
                    }
                }
            }

            console.log(`RSC_Theme FloodLight工作表中未找到主题: ${themeName}`);
            return null;

        } catch (error) {
            console.error('从RSC_Theme FloodLight工作表读取字段时出错:', error);
            return null;
        }
    }

    /**
     * 从RSC_Theme FloodLight工作表读取第一个主题的字段值（用于新建主题）
     * @param {string} floodLightField - FloodLight字段名称
     * @returns {string|null} FloodLight字段值或null
     */
    function findFloodLightValueFromRSCThemeFloodLightFirstTheme(floodLightField) {
        console.log(`=== 从RSC_Theme FloodLight工作表读取第一个主题的字段值: ${floodLightField} ===`);

        try {
            if (!rscAllSheetsData || !rscAllSheetsData['FloodLight']) {
                console.warn('RSC_Theme FloodLight数据未加载');
                return null;
            }

            const floodLightData = rscAllSheetsData['FloodLight'];

            // 检查数据是否足够（需要至少6行：表头+5行元数据+第一个主题）
            if (!floodLightData || floodLightData.length <= 5) {
                console.warn('RSC_Theme FloodLight工作表数据不足（需要至少6行）');
                return null;
            }

            // 查找字段列索引
            const headerRow = floodLightData[0];
            const fieldColumnIndex = headerRow.findIndex(col => col === floodLightField);

            if (fieldColumnIndex === -1) {
                console.log(`RSC_Theme FloodLight工作表中未找到字段: ${floodLightField}`);
                return null;
            }

            // 读取第一个主题（行索引5，第6行）
            const firstThemeRowIndex = 5;
            const firstThemeRow = floodLightData[firstThemeRowIndex];

            if (!firstThemeRow) {
                console.warn(`RSC_Theme FloodLight工作表第一个主题行不存在（行索引: ${firstThemeRowIndex}）`);
                return null;
            }

            const fieldValue = firstThemeRow[fieldColumnIndex];

            if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                console.log(`✅ 从RSC_Theme FloodLight工作表第一个主题找到: ${floodLightField} = ${fieldValue}`);
                return fieldValue.toString();
            } else {
                console.log(`RSC_Theme FloodLight工作表第一个主题中未找到 ${floodLightField} 字段值`);
                return null;
            }

        } catch (error) {
            console.error('从RSC_Theme FloodLight工作表读取第一个主题的字段时出错:', error);
            return null;
        }
    }

    /**
     * 直接映射模式：FloodLight字段条件读取
     * @param {string} floodLightField - FloodLight字段名称（如Color, TippingPoint, Strength等）
     * @param {boolean} isNewTheme - 是否为新建主题
     * @param {string} themeName - 主题名称
     * @returns {string|null} FloodLight字段值或null
     */
    function findFloodLightValueDirect(floodLightField, isNewTheme = false, themeName = '') {
        console.log(`=== 直接映射模式查找FloodLight字段: ${floodLightField} ===`);
        console.log(`主题类型: ${isNewTheme ? '新建主题' : '更新现有主题'}, 主题名称: ${themeName}`);

        if (!sourceData || !sourceData.workbook) {
            console.warn('源数据未加载，无法进行FloodLight字段条件读取');
            return null;
        }

        // 解析Status工作表获取FloodLight状态
        const statusInfo = parseStatusSheet(sourceData);
        console.log('FloodLight状态信息:', {
            hasFloodLightField: statusInfo.hasFloodLightField,
            floodLightStatus: statusInfo.floodLightStatus,
            isFloodLightValid: statusInfo.isFloodLightValid
        });

        if (!statusInfo.hasFloodLightField) {
            console.warn('Status工作表中没有FloodLight字段，根据主题类型处理');

            if (isNewTheme) {
                // ✅ 新建主题模式：从RSC_Theme FloodLight工作表读取第一个主题的数据
                console.log('新建主题且无FloodLight字段，从RSC_Theme FloodLight工作表读取第一个主题的字段值');
                const rscFloodLightValue = findFloodLightValueFromRSCThemeFloodLightFirstTheme(floodLightField);
                if (rscFloodLightValue) {
                    console.log(`✅ 从RSC_Theme FloodLight工作表第一个主题找到: ${floodLightField} = ${rscFloodLightValue}`);
                    return rscFloodLightValue;
                }
            } else if (themeName) {
                // 更新现有主题模式：直接从RSC_Theme FloodLight工作表读取
                console.log('更新现有主题且无FloodLight字段，直接从RSC_Theme FloodLight工作表读取字段值');
                const rscFloodLightValue = findFloodLightValueFromRSCThemeFloodLight(floodLightField, themeName);
                if (rscFloodLightValue) {
                    console.log(`✅ 从RSC_Theme FloodLight工作表找到: ${floodLightField} = ${rscFloodLightValue}`);
                    return rscFloodLightValue;
                }
            }

            console.log('无FloodLight字段且无法从RSC_Theme读取，返回null');
            return null;
        }

        if (statusInfo.isFloodLightValid) {
            // FloodLight状态为有效(1)
            console.log('FloodLight状态有效，尝试从源数据FloodLight工作表读取');

            const sourceFloodLightValue = findFloodLightValueFromSourceFloodLight(floodLightField);
            if (sourceFloodLightValue !== null && sourceFloodLightValue !== undefined && sourceFloodLightValue !== '') {
                console.log(`✅ 从源数据FloodLight工作表找到: ${floodLightField} = ${sourceFloodLightValue}`);
                return sourceFloodLightValue;
            }

            // 🔧 特殊处理：如果是IsOn字段且FloodLight状态为1，但源数据中没有IsOn字段，默认返回1
            if (floodLightField === 'IsOn') {
                console.log('⚠️ 源数据FloodLight工作表未找到IsOn字段，但FloodLight状态为1，默认返回1（开启）');
                return '1';
            }

            // 如果从源数据FloodLight工作表读取字段值没有找到字段，回退到RSC_Theme FloodLight工作表
            if (isNewTheme) {
                // ✅ 新建主题模式：回退到RSC_Theme FloodLight工作表的第一个主题
                console.log('新建主题且源数据FloodLight工作表未找到字段，回退到RSC_Theme FloodLight工作表的第一个主题查找');
                const rscFloodLightValue = findFloodLightValueFromRSCThemeFloodLightFirstTheme(floodLightField);
                if (rscFloodLightValue) {
                    console.log(`✅ 从RSC_Theme FloodLight工作表第一个主题找到: ${floodLightField} = ${rscFloodLightValue}`);
                    return rscFloodLightValue;
                }
            } else if (themeName) {
                console.log('源数据FloodLight工作表未找到字段，回退到RSC_Theme FloodLight工作表查找');
                const rscFloodLightValue = findFloodLightValueFromRSCThemeFloodLight(floodLightField, themeName);
                if (rscFloodLightValue) {
                    console.log(`✅ 从RSC_Theme FloodLight工作表找到: ${floodLightField} = ${rscFloodLightValue}`);
                    return rscFloodLightValue;
                }
            }

            console.log('FloodLight状态有效但未找到字段值，返回null');
            return null;
        } else {
            // FloodLight状态为无效(0)
            console.log('FloodLight状态无效，忽略源数据FloodLight工作表');

            if (isNewTheme) {
                // ✅ 新建主题模式：从RSC_Theme FloodLight工作表读取第一个主题的数据
                console.log('新建主题且FloodLight状态无效，从RSC_Theme FloodLight工作表读取第一个主题的字段值');
                const rscFloodLightValue = findFloodLightValueFromRSCThemeFloodLightFirstTheme(floodLightField);
                if (rscFloodLightValue) {
                    console.log(`✅ 从RSC_Theme FloodLight工作表第一个主题找到: ${floodLightField} = ${rscFloodLightValue}`);
                    return rscFloodLightValue;
                }
            } else if (themeName) {
                // 更新现有主题模式：直接从RSC_Theme FloodLight工作表读取
                console.log('直接从RSC_Theme FloodLight工作表读取FloodLight字段值');
                const rscFloodLightValue = findFloodLightValueFromRSCThemeFloodLight(floodLightField, themeName);
                if (rscFloodLightValue) {
                    console.log(`✅ 从RSC_Theme FloodLight工作表找到: ${floodLightField} = ${rscFloodLightValue}`);
                    return rscFloodLightValue;
                }
            }

            console.log('FloodLight状态无效且无法从RSC_Theme读取，返回null');
            return null;
        }
    }

    /**
     * 从源数据Custom_Ground_Color工作表读取字段值
     * @param {string} fieldName - 字段名称
     * @returns {string|null} 字段值或null
     */
    function findCustomGroundColorValueFromSourceCustomGroundColor(fieldName) {
        console.log(`=== 从源数据Custom_Ground_Color工作表查找字段: ${fieldName} ===`);

        if (!sourceData || !sourceData.workbook) {
            console.log('源数据不可用');
            return null;
        }

        try {
            const customGroundColorSheet = sourceData.workbook.Sheets['Custom_Ground_Color'];
            if (!customGroundColorSheet) {
                console.log('源数据中未找到Custom_Ground_Color工作表');
                return null;
            }

            const customGroundColorData = XLSX.utils.sheet_to_json(customGroundColorSheet, {
                header: 1,
                defval: '',
                raw: false
            });

            if (!customGroundColorData || customGroundColorData.length < 2) {
                console.log('源数据Custom_Ground_Color工作表数据不足');
                return null;
            }

            const headerRow = customGroundColorData[0];
            const dataRow = customGroundColorData[1]; // 第二行数据

            const fieldColumnIndex = headerRow.findIndex(col => col === fieldName);
            if (fieldColumnIndex === -1) {
                console.log(`源数据Custom_Ground_Color工作表中未找到字段: ${fieldName}`);
                return null;
            }

            const fieldValue = dataRow[fieldColumnIndex];
            if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                console.log(`✅ 从源数据Custom_Ground_Color工作表找到: ${fieldName} = ${fieldValue}`);
                return fieldValue.toString();
            } else {
                console.log(`源数据Custom_Ground_Color工作表字段 ${fieldName} 值为空`);
                return null;
            }

        } catch (error) {
            console.error(`从源数据Custom_Ground_Color工作表读取字段 ${fieldName} 时出错:`, error);
            return null;
        }
    }

    /**
     * 从UGCTheme Custom_Ground_Color工作表读取字段值
     * @param {string} fieldName - 字段名称
     * @param {string} themeName - 主题名称
     * @returns {string|null} 字段值或null
     */
    function findCustomGroundColorValueFromUGCThemeCustomGroundColor(fieldName, themeName) {
        console.log(`=== 从UGCTheme Custom_Ground_Color工作表查找字段: ${fieldName} (主题: ${themeName}) ===`);

        if (!ugcAllSheetsData || !ugcAllSheetsData['Custom_Ground_Color']) {
            console.log('UGCTheme Custom_Ground_Color数据未加载');
            return null;
        }

        try {
            const customGroundColorData = ugcAllSheetsData['Custom_Ground_Color'];
            if (!customGroundColorData || customGroundColorData.length < 2) {
                console.log('UGCTheme Custom_Ground_Color工作表数据不足');
                return null;
            }

            const headerRow = customGroundColorData[0];
            const fieldColumnIndex = headerRow.findIndex(col => col === fieldName);
            if (fieldColumnIndex === -1) {
                console.log(`UGCTheme Custom_Ground_Color工作表中未找到字段: ${fieldName}`);
                return null;
            }

            // 查找notes列
            const notesColumnIndex = headerRow.findIndex(col => col === 'notes');
            if (notesColumnIndex === -1) {
                console.log('UGCTheme Custom_Ground_Color工作表中未找到notes列');
                return null;
            }

            // 查找主题对应的行
            for (let i = 1; i < customGroundColorData.length; i++) {
                const row = customGroundColorData[i];
                if (row[notesColumnIndex] === themeName) {
                    const fieldValue = row[fieldColumnIndex];
                    if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                        console.log(`✅ 从UGCTheme Custom_Ground_Color工作表找到: ${fieldName} = ${fieldValue} (主题: ${themeName})`);
                        return fieldValue.toString();
                    } else {
                        console.log(`UGCTheme Custom_Ground_Color工作表中主题 ${themeName} 的字段 ${fieldName} 值为空`);
                        return null;
                    }
                }
            }

            console.log(`UGCTheme Custom_Ground_Color工作表中未找到主题: ${themeName}`);
            return null;

        } catch (error) {
            console.error(`从UGCTheme Custom_Ground_Color工作表读取字段 ${fieldName} 时出错:`, error);
            return null;
        }
    }

    /**
     * Custom_Ground_Color字段条件读取逻辑
     * @param {string} fieldName - 字段名称
     * @param {boolean} isNewTheme - 是否为新建主题
     * @param {string} themeName - 主题名称
     * @returns {string|null} 字段值或null
     */
    function findCustomGroundColorValueDirect(fieldName, isNewTheme = false, themeName = '') {
        console.log(`=== 直接映射查找Custom_Ground_Color字段值: ${fieldName} (新主题: ${isNewTheme}, 主题名: ${themeName}) ===`);

        if (!sourceData || !sourceData.workbook) {
            console.warn('源数据不可用');
            return null;
        }

        // 解析Status工作表状态
        const statusInfo = parseStatusSheet(sourceData);
        console.log('Status状态信息:', statusInfo);

        if (!statusInfo.hasCustomGroundColorField) {
            console.log('Status工作表中没有Custom_Ground_Color字段');
            // 没有Custom_Ground_Color字段：更新现有主题时从UGCTheme读取，新建主题返回null
            if (isNewTheme) {
                console.log('新建主题且无Custom_Ground_Color字段，返回null');
                return null;
            } else if (themeName) {
                console.log('更新现有主题，从UGCTheme Custom_Ground_Color工作表读取');
                return findCustomGroundColorValueFromUGCThemeCustomGroundColor(fieldName, themeName);
            }
            return null;
        }

        const customGroundColorStatus = statusInfo.customGroundColorStatus;
        console.log(`Custom_Ground_Color状态: ${customGroundColorStatus}`);

        if (customGroundColorStatus === 1) {
            console.log('Custom_Ground_Color状态为1（有效），优先从源数据读取');
            // 优先从源数据读取
            const sourceValue = findCustomGroundColorValueFromSourceCustomGroundColor(fieldName);
            if (sourceValue !== null) {
                return sourceValue;
            }

            console.log('源数据中未找到，回退到UGCTheme Custom_Ground_Color工作表');
            // 回退到UGCTheme
            if (themeName) {
                return findCustomGroundColorValueFromUGCThemeCustomGroundColor(fieldName, themeName);
            }
            return null;

        } else {
            console.log('Custom_Ground_Color状态为0（无效），忽略源数据，仅从UGCTheme读取');
            // 状态为0：忽略源数据，仅从UGCTheme读取
            if (themeName) {
                return findCustomGroundColorValueFromUGCThemeCustomGroundColor(fieldName, themeName);
            }
            return null;
        }
    }

    /**
     * 从源数据Custom_Fragile_Color工作表读取字段值
     * @param {string} fieldName - 字段名称
     * @returns {string|null} 字段值或null
     */
    function findCustomFragileColorValueFromSourceCustomFragileColor(fieldName) {
        console.log(`=== 从源数据Custom_Fragile_Color工作表查找字段: ${fieldName} ===`);

        if (!sourceData || !sourceData.workbook) {
            console.log('源数据不可用');
            return null;
        }

        try {
            const customFragileColorSheet = sourceData.workbook.Sheets['Custom_Fragile_Color'];
            if (!customFragileColorSheet) {
                console.log('源数据中未找到Custom_Fragile_Color工作表');
                return null;
            }

            const customFragileColorData = XLSX.utils.sheet_to_json(customFragileColorSheet, {
                header: 1,
                defval: '',
                raw: false
            });

            if (!customFragileColorData || customFragileColorData.length < 2) {
                console.log('源数据Custom_Fragile_Color工作表数据不足');
                return null;
            }

            const headerRow = customFragileColorData[0];
            const dataRow = customFragileColorData[1];

            const fieldColumnIndex = headerRow.findIndex(col => col === fieldName);
            if (fieldColumnIndex === -1) {
                console.log(`源数据Custom_Fragile_Color工作表中未找到字段: ${fieldName}`);
                return null;
            }

            const fieldValue = dataRow[fieldColumnIndex];
            if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                console.log(`✅ 从源数据Custom_Fragile_Color工作表找到: ${fieldName} = ${fieldValue}`);
                return fieldValue.toString();
            } else {
                console.log(`源数据Custom_Fragile_Color工作表字段 ${fieldName} 值为空`);
                return null;
            }

        } catch (error) {
            console.error(`从源数据Custom_Fragile_Color工作表读取字段 ${fieldName} 时出错:`, error);
            return null;
        }
    }

    /**
     * 从UGCTheme Custom_Fragile_Color工作表读取字段值
     * @param {string} fieldName - 字段名称
     * @param {string} themeName - 主题名称
     * @returns {string|null} 字段值或null
     */
    function findCustomFragileColorValueFromUGCThemeCustomFragileColor(fieldName, themeName) {
        console.log(`=== 从UGCTheme Custom_Fragile_Color工作表查找字段: ${fieldName} (主题: ${themeName}) ===`);

        if (!ugcAllSheetsData || !ugcAllSheetsData['Custom_Fragile_Color']) {
            console.log('UGCTheme Custom_Fragile_Color数据未加载');
            return null;
        }

        try {
            const customFragileColorData = ugcAllSheetsData['Custom_Fragile_Color'];
            if (!customFragileColorData || customFragileColorData.length < 2) {
                console.log('UGCTheme Custom_Fragile_Color工作表数据不足');
                return null;
            }

            const headerRow = customFragileColorData[0];
            const fieldColumnIndex = headerRow.findIndex(col => col === fieldName);
            if (fieldColumnIndex === -1) {
                console.log(`UGCTheme Custom_Fragile_Color工作表中未找到字段: ${fieldName}`);
                return null;
            }

            const notesColumnIndex = headerRow.findIndex(col => col === 'notes');
            if (notesColumnIndex === -1) {
                console.log('UGCTheme Custom_Fragile_Color工作表中未找到notes列');
                return null;
            }

            for (let i = 1; i < customFragileColorData.length; i++) {
                const row = customFragileColorData[i];
                if (row[notesColumnIndex] === themeName) {
                    const fieldValue = row[fieldColumnIndex];
                    if (fieldValue !== undefined && fieldValue !== null && fieldValue !== '') {
                        console.log(`✅ 从UGCTheme Custom_Fragile_Color工作表找到: ${fieldName} = ${fieldValue} (主题: ${themeName})`);
                        return fieldValue.toString();
                    } else {
                        console.log(`UGCTheme Custom_Fragile_Color工作表中主题 ${themeName} 的字段 ${fieldName} 值为空`);
                        return null;
                    }
                }
            }

            console.log(`UGCTheme Custom_Fragile_Color工作表中未找到主题: ${themeName}`);
            return null;

        } catch (error) {
            console.error(`从UGCTheme Custom_Fragile_Color工作表读取字段 ${fieldName} 时出错:`, error);
            return null;
        }
    }

    /**
     * Custom_Fragile_Color字段条件读取逻辑
     * @param {string} fieldName - 字段名称
     * @param {boolean} isNewTheme - 是否为新建主题
     * @param {string} themeName - 主题名称
     * @returns {string|null} 字段值或null
     */
    function findCustomFragileColorValueDirect(fieldName, isNewTheme = false, themeName = '') {
        console.log(`=== 直接映射查找Custom_Fragile_Color字段值: ${fieldName} (新主题: ${isNewTheme}, 主题名: ${themeName}) ===`);

        if (!sourceData || !sourceData.workbook) {
            console.warn('源数据不可用');
            return null;
        }

        const statusInfo = parseStatusSheet(sourceData);
        console.log('Status状态信息:', statusInfo);

        if (!statusInfo.hasCustomFragileColorField) {
            console.log('Status工作表中没有Custom_Fragile_Color字段');
            if (isNewTheme) {
                console.log('新建主题且无Custom_Fragile_Color字段，返回null');
                return null;
            } else if (themeName) {
                console.log('更新现有主题，从UGCTheme Custom_Fragile_Color工作表读取');
                return findCustomFragileColorValueFromUGCThemeCustomFragileColor(fieldName, themeName);
            }
            return null;
        }

        const customFragileColorStatus = statusInfo.customFragileColorStatus;
        console.log(`Custom_Fragile_Color状态: ${customFragileColorStatus}`);

        if (customFragileColorStatus === 1) {
            console.log('Custom_Fragile_Color状态为1（有效），优先从源数据读取');
            const sourceValue = findCustomFragileColorValueFromSourceCustomFragileColor(fieldName);
            if (sourceValue !== null) {
                return sourceValue;
            }

            console.log('源数据中未找到，回退到UGCTheme Custom_Fragile_Color工作表');
            if (themeName) {
                return findCustomFragileColorValueFromUGCThemeCustomFragileColor(fieldName, themeName);
            }
            return null;

        } else {
            console.log('Custom_Fragile_Color状态为0（无效），忽略源数据，仅从UGCTheme读取');
            if (themeName) {
                return findCustomFragileColorValueFromUGCThemeCustomFragileColor(fieldName, themeName);
            }
            return null;
        }
    }

    /**
     * 从源数据Custom_Fragile_Active_Color工作表读取字段值
     * @param {string} fieldName - 字段名称
     * @returns {string|null} 字段值或null
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
     * @param {string} fieldName - 字段名称
     * @param {string} themeName - 主题名称
     * @returns {string|null} 字段值或null
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
     * @param {string} fieldName - 字段名称
     * @param {boolean} isNewTheme - 是否为新建主题
     * @param {string} themeName - 主题名称
     * @returns {string|null} 字段值或null
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
            if (isNewTheme) {
                console.log('新建主题且无Custom_Fragile_Active_Color字段，返回null');
                return null;
            } else if (themeName) {
                console.log('更新现有主题，从UGCTheme Custom_Fragile_Active_Color工作表读取');
                return findCustomFragileActiveColorValueFromUGCThemeCustomFragileActiveColor(fieldName, themeName);
            }
            return null;
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
     * 从源数据Custom_Jump_Color工作表读取字段值
     * @param {string} fieldName - 字段名称
     * @returns {string|null} 字段值或null
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
     * @param {string} fieldName - 字段名称
     * @param {string} themeName - 主题名称
     * @returns {string|null} 字段值或null
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
     * @param {string} fieldName - 字段名称
     * @param {boolean} isNewTheme - 是否为新建主题
     * @param {string} themeName - 主题名称
     * @returns {string|null} 字段值或null
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
            if (isNewTheme) {
                console.log('新建主题且无Custom_Jump_Color字段，返回null');
                return null;
            } else if (themeName) {
                console.log('更新现有主题，从UGCTheme Custom_Jump_Color工作表读取');
                return findCustomJumpColorValueFromUGCThemeCustomJumpColor(fieldName, themeName);
            }
            return null;
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
     * 从源数据Custom_Jump_Active_Color工作表读取字段值
     * @param {string} fieldName - 字段名称
     * @returns {string|null} 字段值或null
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
     * @param {string} fieldName - 字段名称
     * @param {string} themeName - 主题名称
     * @returns {string|null} 字段值或null
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
     * @param {string} fieldName - 字段名称
     * @param {boolean} isNewTheme - 是否为新建主题
     * @param {string} themeName - 主题名称
     * @returns {string|null} 字段值或null
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
            if (isNewTheme) {
                console.log('新建主题且无Custom_Jump_Active_Color字段，返回null');
                return null;
            } else if (themeName) {
                console.log('更新现有主题，从UGCTheme Custom_Jump_Active_Color工作表读取');
                return findCustomJumpActiveColorValueFromUGCThemeCustomJumpActiveColor(fieldName, themeName);
            }
            return null;
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

    /**
     * 直接映射模式：Light字段条件读取
     * @param {string} lightField - Light字段名称（如Max, Dark, Min等）
     * @param {boolean} isNewTheme - 是否为新建主题
     * @param {string} themeName - 主题名称
     * @returns {string|null} Light字段值或null
     */
    function findLightValueDirect(lightField, isNewTheme = false, themeName = '') {
        console.log(`=== 直接映射查找Light字段值: ${lightField} (新主题: ${isNewTheme}, 主题名: ${themeName}) ===`);

        if (!sourceData || !sourceData.workbook) {
            console.warn('源数据不可用');
            return null;
        }

        // 解析Status工作表状态
        const statusInfo = parseStatusSheet(sourceData);
        console.log('Status状态信息:', statusInfo);

        if (!statusInfo.hasLightField) {
            console.warn('Status工作表中没有Light字段，根据主题类型处理');

            if (isNewTheme) {
                // ✅ 新建主题模式：从RSC_Theme Light工作表读取第一个主题的数据
                console.log('新建主题且无Light字段，从RSC_Theme Light工作表读取第一个主题的字段值');
                const rscLightValue = findLightValueFromRSCThemeLightFirstTheme(lightField);
                if (rscLightValue) {
                    console.log(`✅ 从RSC_Theme Light工作表第一个主题找到: ${lightField} = ${rscLightValue}`);
                    return rscLightValue;
                }
            } else if (themeName) {
                // 更新现有主题模式：直接从RSC_Theme Light工作表读取
                console.log('更新现有主题且无Light字段，直接从RSC_Theme Light工作表读取字段值');
                const rscLightValue = findLightValueFromRSCThemeLight(lightField, themeName);
                if (rscLightValue) {
                    console.log(`✅ 从RSC_Theme Light工作表找到: ${lightField} = ${rscLightValue}`);
                    return rscLightValue;
                }
            }

            // 未找到：返回null，使用默认值
            console.log(`⚠️ 无Light字段，${isNewTheme ? '新建主题' : '现有主题'}未找到Light字段值: ${lightField}`);
            return null;
        }

        const isLightValid = statusInfo.isLightValid;
        console.log(`Light字段状态: ${isLightValid ? '有效(1)' : '无效(0)'}`);

        // 根据Light状态和主题类型决定处理逻辑
        if (isLightValid) {
            // Light状态为有效(1)
            console.log('Light状态有效，优先从源数据Light工作表查找');

            // 优先从源数据Light工作表查找
            const sourceLightValue = findLightValueFromSourceLight(lightField);
            if (sourceLightValue) {
                console.log(`✅ 从源数据Light工作表找到: ${lightField} = ${sourceLightValue}`);
                return sourceLightValue;
            }

            if (isNewTheme) {
                // ✅ 新建主题模式：回退到RSC_Theme Light工作表的第一个主题
                console.log('新建主题且源数据Light工作表未找到，回退到RSC_Theme Light工作表的第一个主题查找');
                const rscLightValue = findLightValueFromRSCThemeLightFirstTheme(lightField);
                if (rscLightValue) {
                    console.log(`✅ 从RSC_Theme Light工作表第一个主题找到: ${lightField} = ${rscLightValue}`);
                    return rscLightValue;
                }
            } else if (themeName) {
                // 更新现有主题模式：回退到RSC_Theme Light工作表
                console.log('源数据Light工作表未找到，回退到RSC_Theme Light工作表查找');
                const rscLightValue = findLightValueFromRSCThemeLight(lightField, themeName);
                if (rscLightValue) {
                    console.log(`✅ 从RSC_Theme Light工作表找到: ${lightField} = ${rscLightValue}`);
                    return rscLightValue;
                }
            }

            console.log(`⚠️ Light状态有效但未找到Light字段值: ${lightField}`);
            return null;

        } else {
            // Light状态为无效(0)
            console.log('Light状态无效，忽略源数据Light工作表');

            if (isNewTheme) {
                // ✅ 新建主题模式：从RSC_Theme Light工作表读取第一个主题的数据
                console.log('新建主题且Light状态无效，从RSC_Theme Light工作表读取第一个主题的字段值');
                const rscLightValue = findLightValueFromRSCThemeLightFirstTheme(lightField);
                if (rscLightValue) {
                    console.log(`✅ 从RSC_Theme Light工作表第一个主题找到: ${lightField} = ${rscLightValue}`);
                    return rscLightValue;
                }
            } else if (themeName) {
                // 更新现有主题模式：直接从RSC_Theme Light工作表读取
                console.log('直接从RSC_Theme Light工作表读取Light字段值');
                const rscLightValue = findLightValueFromRSCThemeLight(lightField, themeName);
                if (rscLightValue) {
                    console.log(`✅ 从RSC_Theme Light工作表找到: ${lightField} = ${rscLightValue}`);
                    return rscLightValue;
                }
            }

            // 未找到：返回null，使用默认值
            console.log(`⚠️ Light状态无效，${isNewTheme ? '新建主题' : '现有主题'}未找到Light字段值: ${lightField}`);
            return null;
        }
    }

    /**
     * 直接映射模式：更新主题颜色数据
     * @param {number} rowIndex - 主题行索引
     * @param {string} themeName - 主题名称
     * @param {boolean} isNewTheme - 是否为新建主题
     * @returns {Object} 更新结果
     */
    function updateThemeColorsDirect(rowIndex, themeName, isNewTheme = false) {
        console.log('=== 开始直接映射模式更新主题颜色数据 ===');
        console.log(`目标行索引: ${rowIndex}, 主题名称: ${themeName}, 是否新主题: ${isNewTheme}`);

        const data = rscThemeData.data;
        const headerRow = data[0];
        const themeRow = data[rowIndex];

        if (!themeRow) {
            throw new Error(`无法找到行索引 ${rowIndex} 对应的主题行数据`);
        }

        const updatedColors = [];
        const summary = {
            total: 0,
            updated: 0,
            notFound: 0,
            errors: []
        };

        console.log('RSC_Theme表头:', headerRow);

        // 识别所有颜色通道列（P开头和G开头的列）
        const colorChannels = headerRow.filter((col) => {
            if (!col || typeof col !== 'string') return false;
            const colName = col.toString().trim().toUpperCase();
            return colName.startsWith('P') || colName.startsWith('G');
        });

        console.log('发现的颜色通道:', colorChannels);

        // 直接映射处理每个颜色通道
        colorChannels.forEach((channel, index) => {
            const columnIndex = headerRow.findIndex(col => col === channel);

            console.log(`\n处理直接映射 ${index + 1}/${colorChannels.length}: ${channel}`);

            summary.total++;

            try {
                // 直接查找对应的颜色值，传递isNewTheme和themeName参数
                const colorValue = findColorValueDirect(channel, isNewTheme, themeName);

                let finalColorValue = null;
                let isDefault = false;

                if (colorValue && colorValue !== null && colorValue !== undefined && colorValue !== '') {
                    finalColorValue = colorValue;
                    console.log(`✅ 直接映射找到颜色值: ${channel} = ${finalColorValue}`);
                } else {
                    // 使用默认值FFFFFF
                    finalColorValue = 'FFFFFF';
                    isDefault = true;
                    console.log(`⚠️ 直接映射未找到颜色值，使用默认值: ${channel} = ${finalColorValue}`);
                }

                // 更新数据
                if (columnIndex !== -1 && themeRow && columnIndex >= 0 && columnIndex < themeRow.length) {
                    themeRow[columnIndex] = finalColorValue;
                    console.log(`📝 直接映射更新: 行${rowIndex}, 列${columnIndex}(${channel}) = ${finalColorValue}`);

                    // 记录更新结果
                    updatedColors.push({
                        channel: channel,
                        colorCode: channel, // 直接映射模式下，颜色代码就是通道名
                        value: finalColorValue,
                        isDefault: isDefault,
                        rowIndex: rowIndex,
                        columnIndex: columnIndex
                    });

                    if (isDefault) {
                        summary.notFound++;
                    } else {
                        summary.updated++;
                    }
                } else {
                    console.error(`❌ 数据更新失败: 无效的列索引 - 通道:${channel}, 列:${columnIndex}`);
                    summary.errors.push(`无效的列索引: ${channel}`);
                }

            } catch (error) {
                console.error(`处理通道 ${channel} 时出错:`, error);
                summary.errors.push(`处理 ${channel} 时出错: ${error.message}`);
            }
        });

        console.log('\n=== 直接映射模式颜色处理完成 ===');
        console.log('处理统计:', summary);
        console.log('成功更新数量:', summary.updated);
        console.log('使用默认值数量:', summary.notFound);
        console.log('错误数量:', summary.errors.length);

        // 验证数据更新结果
        console.log('=== 数据更新验证 ===');
        console.log(`主题行数据 (行${rowIndex}):`, themeRow);

        return {
            success: true,
            updatedColors: updatedColors,
            summary: summary,
            mode: 'direct'
        };
    }

    /**
     * 更新主题颜色数据（JSON映射模式）
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
                    // 设置默认值（蓝色：5C84F1）
                    const defaultValue = '5C84F1';
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
     * 根据Status工作表状态获取需要处理的工作表列表
     * @param {boolean} isNewTheme - 是否为新建主题
     * @returns {Array} 需要处理的工作表名称数组
     */
    function getActiveSheetsByStatus(isNewTheme = false) {
        console.log('=== 开始根据Status工作表状态获取需要处理的工作表列表 ===');
        console.log('是否新建主题:', isNewTheme);

        // 默认的所有可能工作表
        const allPossibleSheets = ['ColorInfo', 'Light', 'FloodLight', 'VolumetricFog'];

        // 🔧 修复：新建主题时，总是处理所有工作表，不受映射模式和Status状态限制
        if (isNewTheme) {
            console.log('新建主题模式，处理所有工作表（不受Status状态限制）');
            return allPossibleSheets;
        }

        // 🔧 修复：只有在直接映射模式下，才严格按照Status工作表状态决定处理哪些工作表
        // 间接映射模式保持原逻辑，处理所有工作表
        if (currentMappingMode !== 'direct') {
            console.log('非直接映射模式，处理所有工作表');
            return allPossibleSheets;
        }

        // 直接映射模式：解析Status工作表状态
        const statusInfo = parseStatusSheet(sourceData);
        console.log('直接映射模式 - Status状态信息:', statusInfo);

        const activeSheets = [];

        // 根据各字段状态决定是否处理对应工作表
        if (statusInfo.hasColorInfoField && statusInfo.colorInfoStatus === 1) {
            activeSheets.push('ColorInfo');
            console.log('✅ ColorInfo状态为1，添加到处理列表');
        } else {
            console.log('⚠️ ColorInfo状态为0或不存在，跳过处理');
        }

        if (statusInfo.hasLightField && statusInfo.lightStatus === 1) {
            activeSheets.push('Light');
            console.log('✅ Light状态为1，添加到处理列表');
        } else {
            console.log('⚠️ Light状态为0或不存在，跳过处理');
        }

        if (statusInfo.hasVolumetricFogField && statusInfo.volumetricFogStatus === 1) {
            activeSheets.push('VolumetricFog');
            console.log('✅ VolumetricFog状态为1，添加到处理列表');
        } else {
            console.log('⚠️ VolumetricFog状态为0或不存在，跳过处理');
        }

        // FloodLight独立状态驱动处理
        if (statusInfo.hasFloodLightField && statusInfo.floodLightStatus === 1) {
            activeSheets.push('FloodLight');
            console.log('✅ FloodLight状态为1，添加到处理列表');
        } else {
            console.log('⚠️ FloodLight状态为0或不存在，跳过处理');
        }

        console.log(`直接映射模式 - 最终需要处理的工作表: [${activeSheets.join(', ')}]`);
        return activeSheets;
    }

    /**
     * 根据Status工作表状态获取需要处理的UGC工作表列表
     * @returns {Array} 需要处理的UGC工作表名称数组
     */
    function getActiveUGCSheetsByStatus() {
        console.log('=== 开始根据Status工作表状态获取需要处理的UGC工作表列表 ===');

        // 默认的所有可能UGC工作表
        const allPossibleUGCSheets = [
            'Custom_Ground_Color',
            'Custom_Fragile_Color',
            'Custom_Fragile_Active_Color',
            'Custom_Jump_Color',
            'Custom_Jump_Active_Color'
        ];

        // 🔧 修复：只有在直接映射模式下，才严格按照Status工作表状态决定处理哪些UGC工作表
        // 间接映射模式保持原逻辑，处理所有UGC工作表
        if (currentMappingMode !== 'direct') {
            console.log('非直接映射模式，处理所有UGC工作表');
            return allPossibleUGCSheets;
        }

        // 直接映射模式：检查源数据是否可用
        if (!sourceData || !sourceData.workbook) {
            console.log('⚠️ 源数据不可用，无法解析Status工作表，返回空列表');
            return [];
        }

        // 直接映射模式：解析Status工作表状态
        const statusInfo = parseStatusSheet(sourceData);
        console.log('直接映射模式 - Status状态信息:', statusInfo);

        const activeUGCSheets = [];

        // 根据各字段状态决定是否处理对应UGC工作表
        if (statusInfo.hasCustomGroundColorField && statusInfo.customGroundColorStatus === 1) {
            activeUGCSheets.push('Custom_Ground_Color');
            console.log('✅ Custom_Ground_Color状态为1，添加到处理列表');
        } else {
            console.log('⚠️ Custom_Ground_Color状态为0或不存在，跳过处理');
        }

        if (statusInfo.hasCustomFragileColorField && statusInfo.customFragileColorStatus === 1) {
            activeUGCSheets.push('Custom_Fragile_Color');
            console.log('✅ Custom_Fragile_Color状态为1，添加到处理列表');
        } else {
            console.log('⚠️ Custom_Fragile_Color状态为0或不存在，跳过处理');
        }

        if (statusInfo.hasCustomFragileActiveColorField && statusInfo.customFragileActiveColorStatus === 1) {
            activeUGCSheets.push('Custom_Fragile_Active_Color');
            console.log('✅ Custom_Fragile_Active_Color状态为1，添加到处理列表');
        } else {
            console.log('⚠️ Custom_Fragile_Active_Color状态为0或不存在，跳过处理');
        }

        if (statusInfo.hasCustomJumpColorField && statusInfo.customJumpColorStatus === 1) {
            activeUGCSheets.push('Custom_Jump_Color');
            console.log('✅ Custom_Jump_Color状态为1，添加到处理列表');
        } else {
            console.log('⚠️ Custom_Jump_Color状态为0或不存在，跳过处理');
        }

        if (statusInfo.hasCustomJumpActiveColorField && statusInfo.customJumpActiveColorStatus === 1) {
            activeUGCSheets.push('Custom_Jump_Active_Color');
            console.log('✅ Custom_Jump_Active_Color状态为1，添加到处理列表');
        } else {
            console.log('⚠️ Custom_Jump_Active_Color状态为0或不存在，跳过处理');
        }

        console.log(`直接映射模式 - 最终需要处理的UGC工作表: [${activeUGCSheets.join(', ')}]`);
        return activeUGCSheets;
    }

    /**
     * 处理RSC_Theme文件中的ColorInfo、Light和FloodLight sheet（新增主题时添加新行）
     * @param {string} themeName - 主题名称
     * @param {boolean} isNewTheme - 是否为新增主题
     * @returns {Object} 处理结果
     */
    function processRSCAdditionalSheets(themeName, isNewTheme) {
        console.log('=== 开始处理RSC_Theme的ColorInfo、Light和FloodLight sheet ===');
        console.log('主题名称:', themeName);
        console.log('是否新增主题:', isNewTheme);

        if (!isNewTheme) {
            console.log('✅ 更新现有主题，总是处理所有UI配置的工作表（ColorInfo、Light、FloodLight、VolumetricFog）');
            console.log('💡 原因：用户在UI上修改的值应该被保存，无论Status状态如何');
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

            // 🔧 修复：根据Status工作表状态获取需要处理的工作表列表
            const targetSheets = getActiveSheetsByStatus(isNewTheme);
            console.log('🎯 根据Status状态确定的目标工作表:', targetSheets);

            if (targetSheets.length === 0) {
                console.log('⚠️ 没有需要处理的工作表，跳过处理');
                return {
                    success: true,
                    action: 'skip_processing',
                    message: 'Status工作表中没有状态为1的字段，跳过工作表处理',
                    processedSheets: []
                };
            }

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
                        const result = addNewRowToSheet(sheetData, themeName, sheetName, isNewTheme);
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

            // 🔧 修复：为了实现"所见即所得"，总是处理所有UI配置的工作表
            // 即使Status状态为0，用户在UI上修改的值也应该被保存
            // 在applyXXXConfigToRow函数中，会根据Status状态决定数据来源
            const targetSheets = ['ColorInfo', 'Light', 'FloodLight', 'VolumetricFog'];
            console.log('🎯 为了实现所见即所得，处理所有UI配置的工作表:', targetSheets);

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
            applyLightConfigToRow(headerRow, existingRow, themeName, false); // 更新现有主题，isNewTheme=false
        } else if (sheetName === 'ColorInfo') {
            applyColorInfoConfigToRow(headerRow, existingRow, themeName, false); // 更新现有主题，isNewTheme=false
        } else if (sheetName === 'FloodLight') {
            applyFloodLightConfigToRow(headerRow, existingRow, themeName, false); // 更新现有主题，isNewTheme=false
        } else if (sheetName === 'VolumetricFog') {
            applyVolumetricFogConfigToRow(headerRow, existingRow, themeName, false); // 更新现有主题，isNewTheme=false
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
     * @param {boolean} isNewTheme - 是否为新建主题
     * @returns {Object} 处理结果
     */
    function addNewRowToSheet(sheetData, themeName, sheetName, isNewTheme = true) {
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

        // 🔧 修复：找到最后一个有效数据行，避免在空行后添加
        let lastValidRowIndex = sheetData.length - 1;
        while (lastValidRowIndex > 0 && (!sheetData[lastValidRowIndex] || sheetData[lastValidRowIndex].every(cell => !cell || cell === ''))) {
            lastValidRowIndex--;
        }

        // 创建新行，复制最后一个有效行的数据作为模板
        const templateRow = sheetData[lastValidRowIndex];
        const newRow = [...templateRow]; // 复制最后一个有效行数据

        console.log(`最后有效行索引: ${lastValidRowIndex}`);
        console.log(`使用第${lastValidRowIndex}行作为模板:`, templateRow);

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
            applyLightConfigToRow(headerRow, newRow, themeName, isNewTheme);
        } else if (sheetName === 'ColorInfo') {
            applyColorInfoConfigToRow(headerRow, newRow, themeName, isNewTheme);
        } else if (sheetName === 'FloodLight') {
            applyFloodLightConfigToRow(headerRow, newRow, themeName, isNewTheme);
        } else if (sheetName === 'VolumetricFog') {
            applyVolumetricFogConfigToRow(headerRow, newRow, themeName, isNewTheme);
        }

        // 🔧 修复：智能添加新行，避免跳空行
        const newRowIndex = lastValidRowIndex + 1;

        // 如果新行索引小于当前数据长度，则替换现有空行；否则添加新行
        if (newRowIndex < sheetData.length) {
            sheetData[newRowIndex] = newRow;
            console.log(`✅ 新行已替换${sheetName}中的空行，索引: ${newRowIndex}`);
        } else {
            sheetData.push(newRow);
            console.log(`✅ 新行已添加到${sheetName}，索引: ${newRowIndex}`);
        }
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
     * @param {string} themeName - 主题名称（可选，用于直接映射模式）
     * @param {boolean} isNewTheme - 是否为新建主题（可选，用于直接映射模式）
     */
    function applyLightConfigToRow(headerRow, newRow, themeName = '', isNewTheme = false) {
        console.log('=== 开始应用Light配置数据到新行 ===');
        console.log(`主题名称: ${themeName}, 是否新建主题: ${isNewTheme}`);

        try {
            // 检查是否为直接映射模式
            const isDirectMode = currentMappingMode === 'direct';
            console.log(`当前映射模式: ${currentMappingMode}, 是否直接映射: ${isDirectMode}`);

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
                    let value;

                    // 🔧 修复：所见即所得 - 优先使用UI上的值
                    // 获取UI上当前显示的值
                    const lightConfig = getLightConfigData();
                    const uiValue = lightConfig[configKey];

                    if (isDirectMode && themeName) {
                        // 直接映射模式：优先使用UI配置的值（所见即所得）
                        console.log(`直接映射模式：优先使用UI配置值 ${columnName} = ${uiValue}`);
                        value = uiValue;
                    } else {
                        // 非直接映射模式：使用用户配置的数据
                        value = uiValue;
                        console.log(`常规模式使用用户配置: ${columnName} = ${value}`);
                    }

                    newRow[columnIndex] = value.toString();
                    console.log(`Light配置应用: ${columnName} = ${value} (列索引: ${columnIndex})`);
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
     * 应用FloodLight配置数据到新行
     * @param {Array} headerRow - 表头行
     * @param {Array} newRow - 新行数据
     * @param {string} themeName - 主题名称（可选，用于直接映射模式）
     * @param {boolean} isNewTheme - 是否为新建主题（可选，用于直接映射模式）
     */
    function applyFloodLightConfigToRow(headerRow, newRow, themeName = '', isNewTheme = false) {
        console.log('=== 开始应用FloodLight配置数据到新行 ===');
        console.log(`主题名称: ${themeName}, 是否新建主题: ${isNewTheme}`);

        try {
            // 检测映射模式
            const isDirectMode = currentMappingMode === 'direct';
            console.log(`当前映射模式: ${currentMappingMode}, 是否直接映射: ${isDirectMode}`);

            // 检查Status工作表中FloodLight状态
            let floodLightStatusFromStatus = 0;
            if (isDirectMode && sourceData && sourceData.workbook) {
                const statusInfo = parseStatusSheet(sourceData);
                floodLightStatusFromStatus = statusInfo.floodLightStatus;
                console.log(`Status工作表FloodLight状态: ${floodLightStatusFromStatus}`);
            }

            // 定义UI配置字段（这些字段有对应的UI输入控件）
            const uiConfiguredFields = {
                'Color': 'Color',
                'TippingPoint': 'TippingPoint',
                'Strength': 'Strength',
                'IsOn': 'IsOn',
                'JumpActiveIsLightOn': 'JumpActiveIsLightOn',
                'LightStrength': 'LightStrength'
            };

            // 系统字段（跳过处理）
            const systemFields = ['id', 'notes'];

            // 动态处理所有字段
            headerRow.forEach((columnName, columnIndex) => {
                // 跳过系统字段
                if (systemFields.includes(columnName)) {
                    console.log(`跳过系统字段: ${columnName}`);
                    return;
                }

                let value = '';

                // 判断是否为UI配置字段
                if (uiConfiguredFields[columnName]) {
                    // UI配置字段：使用现有逻辑
                    const configKey = uiConfiguredFields[columnName];
                    console.log(`处理UI配置字段: ${columnName} -> ${configKey}`);

                    // 🔧 修复：所见即所得 - 优先使用UI上的值
                    // 获取UI上当前显示的值
                    const floodLightConfig = getFloodLightConfigData();
                    const uiValue = floodLightConfig[configKey] || '0';

                    if (isDirectMode && themeName) {
                        // 直接映射模式：优先使用UI配置的值（所见即所得）
                        // 特殊处理IsOn字段：如果Status工作表中FloodLight状态为1，则自动设置为1
                        if (columnName === 'IsOn' && floodLightStatusFromStatus === 1) {
                            value = '1';
                            console.log(`✅ Status工作表FloodLight状态为1，自动设置IsOn: ${columnName} = ${value}`);
                        } else {
                            value = uiValue;
                            console.log(`✅ 直接映射模式使用UI配置值: ${columnName} = ${value}`);
                        }
                    } else {
                        // 非直接映射模式：使用UI配置的数据
                        value = uiValue;
                        console.log(`非直接映射模式，使用UI配置: ${columnName} = ${value}`);
                    }
                } else {
                    // 非UI配置的字段：使用条件读取逻辑获取正确值
                    console.log(`处理非UI配置字段: ${columnName}`);

                    if (isDirectMode && themeName) {
                        // 直接映射模式：使用条件读取逻辑
                        const directValue = findFloodLightValueDirect(columnName, isNewTheme, themeName);

                        if (directValue !== null && directValue !== undefined && directValue !== '') {
                            value = directValue;
                            console.log(`✅ 直接映射找到非UI字段值: ${columnName} = ${value}`);
                        } else {
                            // 🔧 修复：优化默认值处理逻辑
                            if (!isNewTheme && themeName) {
                                // 更新现有主题：从RSC_Theme获取现有值
                                const rscValue = findFloodLightValueFromRSCThemeFloodLight(columnName, themeName);
                                if (rscValue !== null && rscValue !== undefined && rscValue !== '') {
                                    value = rscValue;
                                    console.log(`从RSC_Theme获取现有值: ${columnName} = ${value}`);
                                } else {
                                    // 如果RSC_Theme也没有值，保持模板行的原有值
                                    value = newRow[columnIndex] || '';
                                    console.log(`保持模板行原有值: ${columnName} = ${value}`);
                                }
                            } else {
                                // 新建主题：保持模板行的原有值，不强制设置为'0'
                                value = newRow[columnIndex] || '';
                                console.log(`新建主题保持模板行值: ${columnName} = ${value}`);
                            }
                        }
                    } else {
                        // 非直接映射模式：从RSC_Theme获取最后一个主题的值
                        const defaultConfig = getLastThemeFloodLightConfig();
                        value = defaultConfig[columnName] || newRow[columnIndex] || '';
                        console.log(`非直接映射模式，使用默认配置: ${columnName} = ${value}`);
                    }
                }

                // 应用值到新行
                newRow[columnIndex] = value;
                console.log(`  设置 ${columnName} (索引${columnIndex}) = ${value}`);
            });

            console.log('✅ FloodLight配置数据应用完成');
        } catch (error) {
            console.error('❌ 应用FloodLight配置数据时出错:', error);
        }
    }

    /**
     * 应用VolumetricFog配置数据到新行
     * @param {Array} headerRow - 表头行
     * @param {Array} newRow - 新行数据
     * @param {string} themeName - 主题名称（可选，用于直接映射模式）
     * @param {boolean} isNewTheme - 是否为新建主题（可选，用于直接映射模式）
     */
    function applyVolumetricFogConfigToRow(headerRow, newRow, themeName = '', isNewTheme = false) {
        console.log('=== 开始应用VolumetricFog配置数据到新行 ===');
        console.log(`主题名称: ${themeName}, 是否新建主题: ${isNewTheme}`);

        try {
            // 检查是否为直接映射模式
            const isDirectMode = currentMappingMode === 'direct';
            console.log(`当前映射模式: ${currentMappingMode}, 是否直接映射: ${isDirectMode}`);

            // 检查Status工作表中VolumetricFog状态
            let volumetricFogStatusFromStatus = 0;
            if (isDirectMode && sourceData && sourceData.workbook) {
                const statusInfo = parseStatusSheet(sourceData);
                volumetricFogStatusFromStatus = statusInfo.volumetricFogStatus;
                console.log(`Status工作表VolumetricFog状态: ${volumetricFogStatusFromStatus}`);
            }

            // UI配置的字段（有UI界面配置）
            const uiConfiguredFields = {
                'Color': 'Color',
                'X': 'X',
                'Y': 'Y',
                'Z': 'Z',
                'Density': 'Density',
                'Rotate': 'Rotate',
                'IsOn': 'IsOn'
            };

            // 跳过的系统字段
            const systemFields = ['id', 'notes'];

            // 动态处理所有字段
            headerRow.forEach((columnName, columnIndex) => {
                if (systemFields.includes(columnName)) {
                    return; // 跳过系统字段
                }

                let value = '';

                if (uiConfiguredFields[columnName]) {
                    // UI配置字段：使用现有逻辑
                    const configKey = uiConfiguredFields[columnName];

                    // 🔧 修复：所见即所得 - 优先使用UI上的值
                    // 获取UI上当前显示的值
                    const volumetricFogConfig = getVolumetricFogConfigData();
                    const uiValue = volumetricFogConfig[configKey] || '0';

                    if (isDirectMode && themeName) {
                        // 直接映射模式：优先使用UI配置的值（所见即所得）
                        // 特殊处理IsOn字段：如果Status工作表中VolumetricFog状态为1，则自动设置为1
                        if (columnName === 'IsOn' && volumetricFogStatusFromStatus === 1) {
                            value = '1';
                            console.log(`✅ Status工作表VolumetricFog状态为1，自动设置IsOn: ${columnName} = ${value}`);
                        } else {
                            value = uiValue;
                            console.log(`✅ 直接映射模式使用UI配置值: ${columnName} = ${value}`);
                        }
                    } else {
                        // 非直接映射模式：使用UI配置
                        value = uiValue;
                        console.log(`非直接映射模式，使用UI配置: ${columnName} = ${value}`);
                    }
                } else {
                    // 非UI配置的字段：使用条件读取逻辑获取正确值
                    console.log(`处理非UI配置字段: ${columnName}`);

                    if (isDirectMode && themeName) {
                        // 直接映射模式：使用条件读取逻辑
                        const directValue = findVolumetricFogValueDirect(columnName, isNewTheme, themeName);

                        if (directValue !== null && directValue !== undefined && directValue !== '') {
                            value = directValue;
                            console.log(`✅ 直接映射找到非UI字段值: ${columnName} = ${value}`);
                        } else {
                            // 🔧 修复：优化默认值处理逻辑
                            if (!isNewTheme && themeName) {
                                // 更新现有主题：从RSC_Theme获取现有值
                                const rscValue = findVolumetricFogValueFromRSCThemeVolumetricFog(columnName, themeName);
                                if (rscValue !== null && rscValue !== undefined && rscValue !== '') {
                                    value = rscValue;
                                    console.log(`从RSC_Theme获取现有值: ${columnName} = ${value}`);
                                } else {
                                    // 如果RSC_Theme也没有值，保持模板行的原有值
                                    value = newRow[columnIndex] || '';
                                    console.log(`保持模板行原有值: ${columnName} = ${value}`);
                                }
                            } else {
                                // 新建主题：保持模板行的原有值，不强制设置为'0'
                                value = newRow[columnIndex] || '';
                                console.log(`新建主题保持模板行值: ${columnName} = ${value}`);
                            }
                        }
                    } else {
                        // 非直接映射模式：从RSC_Theme获取最后一个主题的值
                        const defaultConfig = getLastThemeVolumetricFogConfig();
                        value = defaultConfig[columnName] || newRow[columnIndex] || '';
                        console.log(`非直接映射模式，使用默认配置: ${columnName} = ${value}`);
                    }
                }

                // 应用值到新行
                newRow[columnIndex] = value;
                console.log(`  设置 ${columnName} (索引${columnIndex}) = ${value}`);
            });

            console.log('✅ VolumetricFog配置数据应用完成');
        } catch (error) {
            console.error('❌ 应用VolumetricFog配置数据时出错:', error);
        }
    }

    /**
     * 应用ColorInfo配置数据到新行
     * @param {Array} headerRow - 表头行
     * @param {Array} newRow - 新行数据
     * @param {string} themeName - 主题名称（可选，用于直接映射模式）
     * @param {boolean} isNewTheme - 是否为新建主题（可选，用于直接映射模式）
     */
    function applyColorInfoConfigToRow(headerRow, newRow, themeName = '', isNewTheme = false) {
        console.log('=== 开始应用ColorInfo配置数据到新行 ===');
        console.log(`主题名称: ${themeName}, 是否新建主题: ${isNewTheme}`);

        try {
            // 检查是否为直接映射模式
            const isDirectMode = currentMappingMode === 'direct';
            console.log(`当前映射模式: ${currentMappingMode}, 是否直接映射: ${isDirectMode}`);

            // UI配置的ColorInfo字段映射（这些字段有UI界面配置）
            const uiConfiguredFields = {
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

            // 跳过的系统字段（不需要处理的字段）
            const systemFields = ['id', 'notes'];

            // 处理所有ColorInfo字段
            headerRow.forEach((columnName, columnIndex) => {
                // 跳过系统字段
                if (systemFields.includes(columnName)) {
                    console.log(`跳过系统字段: ${columnName}`);
                    return;
                }

                let value;

                if (uiConfiguredFields[columnName]) {
                    // UI配置的字段：使用现有逻辑
                    const configKey = uiConfiguredFields[columnName];

                    // 🔧 修复：所见即所得 - 优先使用UI上的值
                    // 获取UI上当前显示的值
                    const colorInfoConfig = getColorInfoConfigData();
                    const uiValue = colorInfoConfig[configKey];

                    if (isDirectMode && themeName) {
                        // 直接映射模式：优先使用UI配置的值（所见即所得）
                        value = uiValue;
                        console.log(`✅ 直接映射模式使用UI配置值: ${columnName} = ${value}`);
                    } else {
                        // 非直接映射模式：使用用户配置的数据
                        value = uiValue;
                        console.log(`常规模式使用用户配置: ${columnName} = ${value}`);
                    }
                } else {
                    // 非UI配置的字段：使用条件读取逻辑获取正确值
                    console.log(`处理非UI配置字段: ${columnName}`);

                    if (isDirectMode && themeName) {
                        // 直接映射模式：使用条件读取逻辑
                        const directValue = findColorInfoValueDirect(columnName, isNewTheme, themeName);

                        if (directValue !== null && directValue !== undefined && directValue !== '') {
                            value = directValue;
                            console.log(`✅ 直接映射找到非UI字段值: ${columnName} = ${value}`);
                        } else {
                            // 从RSC_Theme获取默认值
                            if (!isNewTheme && themeName) {
                                const rscValue = findColorInfoValueFromRSCThemeColorInfo(columnName, themeName);
                                if (rscValue !== null && rscValue !== undefined && rscValue !== '') {
                                    value = rscValue;
                                    console.log(`✅ 从RSC_Theme获取非UI字段默认值: ${columnName} = ${value}`);
                                } else {
                                    value = '0'; // 最终默认值
                                    console.log(`⚠️ 使用最终默认值: ${columnName} = ${value}`);
                                }
                            } else {
                                value = '0'; // 新建主题的默认值
                                console.log(`⚠️ 新建主题使用默认值: ${columnName} = ${value}`);
                            }
                        }
                    } else {
                        // 非直接映射模式：从RSC_Theme获取最后一个主题的值
                        const defaultConfig = getLastThemeColorInfoConfig();
                        // 尝试从默认配置中获取，如果没有则使用0
                        value = defaultConfig[columnName] || '0';
                        console.log(`非直接映射模式，非UI字段使用默认配置: ${columnName} = ${value}`);
                    }
                }

                // 应用值到新行
                if (value !== undefined && value !== null) {
                    newRow[columnIndex] = value.toString();
                    console.log(`ColorInfo配置应用: ${columnName} = ${value} (列索引: ${columnIndex})`);
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
    function applyUGCFieldSettings(sheetName, headerRow, newRow, themeName = '', isNewTheme = false) {
        console.log(`\n=== 应用UGC字段设置到Sheet ${sheetName} (主题: ${themeName}, 新主题: ${isNewTheme}) ===`);

        // 检查是否为直接映射模式
        const isDirectMode = currentMappingMode === 'direct';
        console.log(`当前映射模式: ${currentMappingMode}, 直接映射: ${isDirectMode}`);

        // 定义条件读取函数映射
        const conditionalReadFunctions = {
            'Custom_Ground_Color': findCustomGroundColorValueDirect,
            'Custom_Fragile_Color': findCustomFragileColorValueDirect,
            'Custom_Fragile_Active_Color': findCustomFragileActiveColorValueDirect,
            'Custom_Jump_Color': findCustomJumpColorValueDirect,
            'Custom_Jump_Active_Color': findCustomJumpActiveColorValueDirect
        };

        // 获取UI配置数据作为回退
        const ugcConfig = getUGCConfigData();

        // 定义UI配置字段映射（仅作为回退）
        const uiConfigMapping = {
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

        console.log(`\n📊 开始处理Sheet ${sheetName} 的所有字段 (共${headerRow.length}个字段)`);
        console.log(`表头:`, headerRow);

        // 遍历工作表中的所有字段（动态字段处理）
        headerRow.forEach((columnName, columnIndex) => {
            // 跳过空列名
            if (!columnName || columnName.toString().trim() === '') {
                return;
            }

            // 跳过系统字段（id、LevelName等已在其他地方处理）
            const systemFields = ['id', 'LevelName', 'levelname', 'LevelName_ID', 'levelname_id',
                                  'Level_id', 'level_id', 'LevelId', 'Level_show_id', 'Level_show_bg_ID'];
            if (systemFields.some(sf => columnName.toLowerCase().includes(sf.toLowerCase()))) {
                console.log(`⏭️ 跳过系统字段: ${columnName}`);
                return;
            }

            let finalValue = newRow[columnIndex]; // 默认保留模板行的值

            // 如果是直接映射模式且有条件读取函数，尝试从源数据读取
            if (isDirectMode && themeName && conditionalReadFunctions[sheetName]) {
                const conditionalReadFunc = conditionalReadFunctions[sheetName];
                const directValue = conditionalReadFunc(columnName, isNewTheme, themeName);

                if (directValue !== null && directValue !== undefined && directValue !== '') {
                    finalValue = directValue;
                    console.log(`✅ [源数据读取] Sheet ${sheetName} 字段 ${columnName}: ${finalValue}`);
                } else {
                    // 条件读取返回空，检查是否有UI配置值
                    const uiConfig = uiConfigMapping[sheetName];
                    if (uiConfig && uiConfig[columnName] !== undefined) {
                        finalValue = uiConfig[columnName];
                        console.log(`📋 [UI配置] Sheet ${sheetName} 字段 ${columnName}: ${finalValue}`);
                    } else {
                        // 保留模板行的值
                        console.log(`🔄 [保留模板] Sheet ${sheetName} 字段 ${columnName}: ${finalValue}`);
                    }
                }
            } else {
                // 非直接映射模式，使用UI配置或保留模板值
                const uiConfig = uiConfigMapping[sheetName];
                if (uiConfig && uiConfig[columnName] !== undefined) {
                    finalValue = uiConfig[columnName];
                    console.log(`📋 [UI配置] Sheet ${sheetName} 字段 ${columnName}: ${finalValue}`);
                } else {
                    console.log(`🔄 [保留模板] Sheet ${sheetName} 字段 ${columnName}: ${finalValue}`);
                }
            }

            // 设置最终值
            newRow[columnIndex] = finalValue !== null && finalValue !== undefined ? finalValue.toString() : '';
        });

        console.log(`✅ Sheet ${sheetName} 所有字段处理完成`);
        console.log(`最终行数据:`, newRow);
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

            // 🔧 修复：为了实现"所见即所得"，总是处理所有UI配置的UGC工作表
            // 即使Status状态为0，用户在UI上修改的值也应该被保存
            const targetUGCSheets = ['Custom_Ground_Color', 'Custom_Fragile_Color', 'Custom_Fragile_Active_Color', 'Custom_Jump_Color', 'Custom_Jump_Active_Color'];
            console.log('🎯 为了实现所见即所得，总是处理所有UI配置的UGC工作表:', targetUGCSheets);

            // 第三步：更新每个相关的sheet（总是处理所有UI配置的工作表）
            targetUGCSheets.forEach(sheetName => {
                console.log(`\n--- ✅ 更新Sheet: ${sheetName} (总是处理所有UI配置的工作表) ---`);

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

                console.log(`Sheet ${sheetName} 更新行 ${targetRowNumber} (共${headerRow.length}个字段)`);
                console.log(`表头:`, headerRow);
                console.log(`更新前的行数据:`, targetRow);

                // 检查是否为直接映射模式
                const isDirectMode = currentMappingMode === 'direct';
                console.log(`当前映射模式: ${currentMappingMode}, 直接映射: ${isDirectMode}`);

                // 定义条件读取函数映射
                const conditionalReadFunctions = {
                    'Custom_Ground_Color': findCustomGroundColorValueDirect,
                    'Custom_Fragile_Color': findCustomFragileColorValueDirect,
                    'Custom_Fragile_Active_Color': findCustomFragileActiveColorValueDirect,
                    'Custom_Jump_Color': findCustomJumpColorValueDirect,
                    'Custom_Jump_Active_Color': findCustomJumpActiveColorValueDirect
                };

                // 定义UI配置字段映射（仅作为回退）
                const uiConfigMapping = {
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

                const updatedFields = [];

                // 遍历工作表中的所有字段（动态字段处理）
                headerRow.forEach((columnName, columnIndex) => {
                    // 跳过空列名
                    if (!columnName || columnName.toString().trim() === '') {
                        return;
                    }

                    // 跳过系统字段（id、LevelName等已在其他地方处理）
                    const systemFields = ['id', 'LevelName', 'levelname', 'LevelName_ID', 'levelname_id',
                                          'Level_id', 'level_id', 'LevelId', 'Level_show_id', 'Level_show_bg_ID'];
                    if (systemFields.some(sf => columnName.toLowerCase().includes(sf.toLowerCase()))) {
                        console.log(`⏭️ 跳过系统字段: ${columnName}`);
                        return;
                    }

                    const oldValue = targetRow[columnIndex];
                    let finalValue = oldValue; // 默认保留原值

                    // 如果是直接映射模式且有条件读取函数，尝试从源数据读取
                    if (isDirectMode && themeName && conditionalReadFunctions[sheetName]) {
                        const conditionalReadFunc = conditionalReadFunctions[sheetName];
                        const directValue = conditionalReadFunc(columnName, false, themeName); // isNewTheme=false

                        if (directValue !== null && directValue !== undefined && directValue !== '') {
                            finalValue = directValue;
                            console.log(`✅ [源数据读取] Sheet ${sheetName} 字段 ${columnName}: ${oldValue} -> ${finalValue}`);
                            updatedFields.push(columnName);
                        } else {
                            // 条件读取返回空，检查是否有UI配置值
                            const uiConfig = uiConfigMapping[sheetName];
                            if (uiConfig && uiConfig[columnName] !== undefined) {
                                finalValue = uiConfig[columnName];
                                console.log(`📋 [UI配置] Sheet ${sheetName} 字段 ${columnName}: ${oldValue} -> ${finalValue}`);
                                updatedFields.push(columnName);
                            } else {
                                // 保留原值
                                console.log(`🔄 [保留原值] Sheet ${sheetName} 字段 ${columnName}: ${finalValue}`);
                            }
                        }
                    } else {
                        // 非直接映射模式，使用UI配置或保留原值
                        const uiConfig = uiConfigMapping[sheetName];
                        if (uiConfig && uiConfig[columnName] !== undefined) {
                            finalValue = uiConfig[columnName];
                            console.log(`📋 [UI配置] Sheet ${sheetName} 字段 ${columnName}: ${oldValue} -> ${finalValue}`);
                            updatedFields.push(columnName);
                        } else {
                            console.log(`🔄 [保留原值] Sheet ${sheetName} 字段 ${columnName}: ${finalValue}`);
                        }
                    }

                    // 设置最终值
                    targetRow[columnIndex] = finalValue !== null && finalValue !== undefined ? finalValue.toString() : '';
                });

                console.log(`✅ Sheet ${sheetName} 所有字段处理完成，更新了 ${updatedFields.length} 个字段`);
                console.log(`更新后的行数据:`, targetRow);

                // 更新worksheet
                const newWorksheet = XLSX.utils.aoa_to_sheet(data);
                workbook.Sheets[sheetName] = newWorksheet;

                processedSheets.push({
                    sheetName: sheetName,
                    updatedRowIndex: targetRowNumber,
                    updatedFields: updatedFields
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

            // 🔧 定义总是需要处理的工作表（新建主题时，这些表总是复制上一行）
            const alwaysProcessSheets = ['Custom_Enemy_Color', 'Custom_Mover_Color', 'Custom_Mover_Auto_Color'];

            // 获取需要处理的UGC工作表列表（根据Status状态）
            const activeUGCSheets = getActiveUGCSheetsByStatus();
            console.log(`根据Status状态，需要处理的UGC工作表: [${activeUGCSheets.join(', ')}]`);
            console.log(`总是处理的工作表（新建主题）: [${alwaysProcessSheets.join(', ')}]`);

            // 对每个需要处理的sheet进行处理
            for (const sheetName of sheetNames) {
                // 🔧 检查当前工作表是否需要处理
                const isActiveUGCSheet = activeUGCSheets.includes(sheetName);
                const isAlwaysProcessSheet = alwaysProcessSheets.includes(sheetName);

                if (!isActiveUGCSheet && !isAlwaysProcessSheet) {
                    console.log(`⚠️ Sheet ${sheetName} 不在需要处理的列表中，跳过处理`);
                    continue;
                }

                if (isAlwaysProcessSheet) {
                    console.log(`✅ 处理sheet: ${sheetName} (总是处理的工作表)`);
                } else {
                    console.log(`✅ 处理sheet: ${sheetName} (Status状态允许)`);
                }

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

                // 🔧 修复：找到最后一个有效数据行，避免复制空行作为模板
                let lastValidRowIndex = data.length - 1;
                while (lastValidRowIndex > 0 && (!data[lastValidRowIndex] || data[lastValidRowIndex].every(cell => !cell || cell === ''))) {
                    lastValidRowIndex--;
                }

                // 创建新行（复制最后一个有效行数据）
                const lastRow = data[lastValidRowIndex];
                const newRow = [...lastRow]; // 复制最后一个有效行
                newRow[idColumnIndex] = newId.toString(); // 设置新的ID

                console.log(`Sheet ${sheetName} 最后有效行索引: ${lastValidRowIndex}`);
                console.log(`Sheet ${sheetName} 使用第${lastValidRowIndex}行作为模板`);

                // 🔧 修复：计算新行应该插入的位置
                const newRowIndex = lastValidRowIndex + 1;

                // 🔧 检查是否为"总是处理"的工作表（只需简单复制上一行）
                if (isAlwaysProcessSheet) {
                    console.log(`Sheet ${sheetName} 是总是处理的工作表，只复制上一行数据，id=${newId}`);

                    // 🔧 修复：智能添加新行，避免跳空行
                    if (newRowIndex < data.length) {
                        data[newRowIndex] = newRow;
                        console.log(`✅ Sheet ${sheetName} 新行已替换空行，索引: ${newRowIndex}`);
                    } else {
                        data.push(newRow);
                        console.log(`✅ Sheet ${sheetName} 新行已添加到末尾，索引: ${newRowIndex}`);
                    }

                    // 更新工作表
                    const newWorksheet = XLSX.utils.aoa_to_sheet(data);
                    workbook.Sheets[sheetName] = newWorksheet;

                    processedSheets.push({
                        sheetName: sheetName,
                        newId: newId,
                        action: 'simple_copy'
                    });

                    console.log(`✅ Sheet ${sheetName} 处理完成（简单复制模式）`);
                    continue; // 跳过后续复杂处理
                }

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

                // 处理UGC特定字段设置（传入themeName和isNewTheme以支持条件读取）
                applyUGCFieldSettings(sheetName, headerRow, newRow, themeName, true);

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

                        // 🔧 修复：智能添加新行到数据（先添加，再进行排序插入）
                        if (newRowIndex < data.length) {
                            data[newRowIndex] = newRow;
                            console.log(`Sheet ${sheetName} 新行已替换空行，索引: ${newRowIndex}`);
                        } else {
                            data.push(newRow);
                            console.log(`Sheet ${sheetName} 新行已添加到末尾，索引: ${newRowIndex}`);
                        }

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
                        // 🔧 修复：智能添加新行到数据
                        if (newRowIndex < data.length) {
                            data[newRowIndex] = newRow;
                            console.log(`Sheet ${sheetName} 新行已替换空行，索引: ${newRowIndex}`);
                        } else {
                            data.push(newRow);
                            console.log(`Sheet ${sheetName} 新行已添加到末尾，索引: ${newRowIndex}`);
                        }
                    }
                } else {
                    // 非Custom_Ground_Color工作表或非同系列主题，使用默认添加方式
                    console.log(`Sheet ${sheetName}: 使用默认添加方式 (同系列: ${smartConfig.similarity.isSimilar})`);
                    // 🔧 修复：智能添加新行到数据
                    if (newRowIndex < data.length) {
                        data[newRowIndex] = newRow;
                        console.log(`Sheet ${sheetName} 新行已替换空行，索引: ${newRowIndex}`);
                    } else {
                        data.push(newRow);
                        console.log(`Sheet ${sheetName} 新行已添加到末尾，索引: ${newRowIndex}`);
                    }
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

        // 🔧 关键修复：确保其他工作表不被意外修改
        // 问题：rscAllSheetsData包含了所有工作表的数据，如果这些数据在之前被修改了，
        // 那么在保存时会把修改后的数据写回到Excel文件，导致不应该被处理的工作表也被修改了
        // 解决方案：只更新主工作表和targetSheets中的工作表，其他工作表从rscOriginalSheetsData中重新读取
        console.log('=== 🔧 开始重置非目标工作表数据 ===');

        // 🔧 修复：为了实现"所见即所得"，总是包含所有UI配置的工作表
        // 即使Status状态为0，用户在UI上修改的值也应该被保存
        const targetSheets = ['ColorInfo', 'Light', 'FloodLight', 'VolumetricFog'];
        console.log('🎯 为了实现所见即所得，总是处理所有UI配置的工作表:', targetSheets);

        // 定义允许修改的工作表列表（主工作表 + 目标工作表）
        const allowedSheets = [originalSheetName, ...targetSheets];
        console.log('✅ 允许修改的工作表列表:', allowedSheets);

        // 遍历所有工作表，重置非目标工作表的数据
        workbook.SheetNames.forEach(sheetName => {
            if (!allowedSheets.includes(sheetName)) {
                // 这个工作表不在允许修改的列表中，需要从原始数据中重新读取
                console.log(`🔄 重置非目标工作表: ${sheetName}`);

                // 🔧 从rscOriginalSheetsData（原始备份数据）中读取原始数据
                if (rscOriginalSheetsData && rscOriginalSheetsData[sheetName]) {
                    const originalSheetData = rscOriginalSheetsData[sheetName];

                    // 重新创建工作表（使用原始备份数据）
                    const resetWorksheet = XLSX.utils.aoa_to_sheet(originalSheetData);

                    // 替换工作表（使用原始数据）
                    workbook.Sheets[sheetName] = resetWorksheet;

                    // 同时更新rscAllSheetsData，确保内存数据一致
                    if (rscAllSheetsData) {
                        rscAllSheetsData[sheetName] = JSON.parse(JSON.stringify(originalSheetData));
                    }

                    console.log(`✅ 已重置工作表 "${sheetName}" 为原始数据，行数: ${originalSheetData.length}`);
                } else {
                    console.warn(`⚠️ 无法找到工作表 "${sheetName}" 的原始备份数据，跳过重置`);
                }
            }
        });

        // 处理目标工作表（严格限制：仅限Light、ColorInfo）
        // 重要约束：不修改RSC_Theme.xls文件中的其他工作表，保持零影响原则
        console.log('=== 开始处理目标工作表 ===');
        console.log(`主工作表名称: ${originalSheetName}`);

        if (rscAllSheetsData) {
            console.log('rscAllSheetsData可用工作表:', Object.keys(rscAllSheetsData));

            targetSheets.forEach(sheetName => {
                console.log(`检查目标工作表: ${sheetName}`);
                console.log(`- 是否与主工作表同名: ${sheetName === originalSheetName}`);
                console.log(`- rscAllSheetsData中是否存在: ${!!rscAllSheetsData[sheetName]}`);

                // 修复：移除与主工作表名称的冲突检查，允许处理所有目标工作表
                // 原条件：sheetName !== originalSheetName && rscAllSheetsData[sheetName]
                // 新条件：只检查数据是否存在
                if (rscAllSheetsData[sheetName]) {
                    console.log(`开始处理目标工作表: ${sheetName}`);

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

                                // 额外验证：检查新增行是否包含用户配置的数据
                                if (sheetName === 'Light') {
                                    const maxIndex = headerRow.findIndex(col => col === 'Max');
                                    const specularColorIndex = headerRow.findIndex(col => col === 'SpecularColor');
                                    console.log(`  Light配置验证 - Max值: ${lastDataRow[maxIndex]}, SpecularColor: ${lastDataRow[specularColorIndex]}`);
                                } else if (sheetName === 'ColorInfo') {
                                    const pickupDiffRIndex = headerRow.findIndex(col => col === 'PickupDiffR');
                                    const fogStartIndex = headerRow.findIndex(col => col === 'FogStart');
                                    console.log(`  ColorInfo配置验证 - PickupDiffR: ${lastDataRow[pickupDiffRIndex]}, FogStart: ${lastDataRow[fogStartIndex]}`);
                                }
                            }
                        }
                    } else {
                        console.warn(`工作表 "${sheetName}" 数据为空，跳过更新`);
                    }
                } else {
                    console.warn(`工作表 "${sheetName}" 在rscAllSheetsData中不存在，跳过处理`);
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
     * @param {Object} allObstacleResult - AllObstacle处理结果
     */
    async function handleFileSave(workbook, themeName, ugcResult, allObstacleResult) {
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

                        // 构建保存状态消息
                        let statusMessages = ['RSC_Theme文件保存成功'];
                        if (ugcMessage) {
                            statusMessages.push(ugcMessage);
                        }
                        if (allObstacleResult && allObstacleResult.success) {
                            statusMessages.push('AllObstacle文件处理成功');
                        } else if (allObstacleResult && allObstacleResult.skipped) {
                            statusMessages.push('AllObstacle文件已跳过');
                        }

                        const message = statusMessages.join('，');
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
     * 处理AllObstacle.xls文件（仅全新系列主题时）
     *
     * 功能说明：
     * - 仅在创建全新系列主题时触发
     * - 在AllObstacle.xls的Info工作表中新增一行数据
     * - 自动计算新的ID和Sort值
     * - 填充主题基础名称和多语言ID
     *
     * 字段填充规则：
     * - id列：现有最大值+1
     * - notes列：主题基础名称（去除数字后缀）
     * - nameID列：用户输入的多语言ID
     * - Sort列：现有最大值+1
     * - isFilter列：固定值1
     *
     * @param {string} themeName - 主题名称
     * @param {number} multiLangId - 多语言ID
     * @returns {Object} 处理结果
     * @returns {boolean} returns.success - 是否成功
     * @returns {boolean} returns.updated - 是否更新了文件
     * @returns {boolean} returns.skipped - 是否跳过处理
     * @returns {string} returns.message - 处理消息
     * @returns {string} returns.reason - 跳过原因（如果跳过）
     * @returns {string} returns.error - 错误信息（如果失败）
     */
    async function processAllObstacle(themeName, multiLangId) {
        console.log('=== 开始处理AllObstacle.xls文件 ===');
        console.log('主题名称:', themeName);
        console.log('多语言ID:', multiLangId);

        try {
            // 检查是否有AllObstacle文件句柄
            if (!folderManager || !folderManager.allObstacleHandle) {
                console.log('AllObstacle.xls文件未找到，跳过处理');
                App.Utils.showStatus('AllObstacle.xls文件未找到，跳过AllObstacle处理', 'info', 3000);
                return {
                    success: false,
                    skipped: true,
                    reason: 'AllObstacle.xls文件未找到，请确保该文件位于Unity项目文件夹中'
                };
            }

            // 检查多语言ID
            if (!multiLangId || isNaN(multiLangId) || multiLangId <= 0) {
                console.log('多语言ID无效，跳过AllObstacle处理');
                return {
                    success: false,
                    skipped: true,
                    reason: '多语言ID无效或未提供'
                };
            }

            // 读取AllObstacle文件
            console.log('读取AllObstacle.xls文件...');
            const allObstacleFileData = await folderManager.loadThemeFileData('allObstacle');

            if (!allObstacleFileData || !allObstacleFileData.workbook) {
                throw new Error('AllObstacle.xls文件读取失败');
            }

            const workbook = allObstacleFileData.workbook;
            const sheetNames = workbook.SheetNames;
            console.log('AllObstacle文件包含的工作表:', sheetNames);

            // 查找Info工作表
            const infoSheetName = sheetNames.find(name =>
                name === 'Info' ||
                name.toLowerCase() === 'info' ||
                name.toLowerCase().includes('info')
            );

            if (!infoSheetName) {
                throw new Error('AllObstacle.xls文件中未找到Info工作表');
            }

            console.log('找到Info工作表:', infoSheetName);
            const infoSheet = workbook.Sheets[infoSheetName];
            const infoData = XLSX.utils.sheet_to_json(infoSheet, { header: 1 });

            if (infoData.length === 0) {
                throw new Error('Info工作表为空');
            }

            // 处理Info工作表数据
            const result = await updateAllObstacleInfoData(infoData, themeName, multiLangId);

            if (result.updated) {
                // 保存更新后的文件
                const updatedSheet = XLSX.utils.aoa_to_sheet(result.data);
                workbook.Sheets[infoSheetName] = updatedSheet;

                // 保存文件
                await saveAllObstacleFileDirectly(workbook);

                console.log('✅ AllObstacle.xls文件处理成功');
                return {
                    success: true,
                    updated: true,
                    message: result.message,
                    workbook: workbook
                };
            } else {
                console.log('AllObstacle.xls文件无需更新');
                return {
                    success: true,
                    updated: false,
                    message: result.message
                };
            }

        } catch (error) {
            console.error('AllObstacle.xls文件处理失败:', error);
            return {
                success: false,
                error: error.message
            };
        }
    }

    /**
     * 更新AllObstacle Info工作表数据
     * @param {Array} infoData - Info工作表数据数组
     * @param {string} themeName - 主题名称
     * @param {number} multiLangId - 多语言ID
     */
    async function updateAllObstacleInfoData(infoData, themeName, multiLangId) {
        console.log('=== 开始更新AllObstacle Info数据 ===');

        if (infoData.length < 6) {
            throw new Error('Info工作表数据不足，至少需要6行（表头在第1行，数据从第6行开始）');
        }

        // AllObstacle.xls的Info工作表结构：表头在第1行，数据从第6行开始
        const headerRow = infoData[0];
        console.log('Info工作表表头:', headerRow);
        console.log('Info工作表总行数:', infoData.length);
        console.log('数据行范围: 第6行到第', infoData.length, '行');

        // 查找必要的列
        const idColumnIndex = headerRow.findIndex(col => col === 'id');
        const notesColumnIndex = headerRow.findIndex(col => col === 'notes');
        const nameIDColumnIndex = headerRow.findIndex(col => col === 'nameID');
        const sortColumnIndex = headerRow.findIndex(col => col === 'Sort');
        const isFilterColumnIndex = headerRow.findIndex(col => col === 'isFilter');

        if (idColumnIndex === -1 || notesColumnIndex === -1 || nameIDColumnIndex === -1 ||
            sortColumnIndex === -1 || isFilterColumnIndex === -1) {
            throw new Error('Info工作表中缺少必要的列：id、notes、nameID、Sort、isFilter');
        }

        // 检查是否已存在该多语言ID（数据从第6行开始，索引为5）
        let existingRowIndex = -1;
        const dataStartRow = 5; // 第6行的索引是5

        console.log(`开始检查多语言ID ${multiLangId} 是否已存在（从第${dataStartRow + 1}行开始检查）...`);

        for (let i = dataStartRow; i < infoData.length; i++) {
            const row = infoData[i];
            if (row && row[nameIDColumnIndex] && parseInt(row[nameIDColumnIndex]) === multiLangId) {
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

        // 获取最大ID和Sort值（从第6行开始扫描）
        const existingIds = [];
        const existingSorts = [];

        console.log(`开始扫描现有ID和Sort值（从第${dataStartRow + 1}行到第${infoData.length}行）...`);

        for (let i = dataStartRow; i < infoData.length; i++) {
            const row = infoData[i];
            if (row && row[idColumnIndex]) {
                const id = parseInt(row[idColumnIndex]);
                if (!isNaN(id)) {
                    existingIds.push(id);
                    console.log(`第${i + 1}行 ID: ${id}`);
                }
            }
            if (row && row[sortColumnIndex]) {
                const sort = parseInt(row[sortColumnIndex]);
                if (!isNaN(sort)) {
                    existingSorts.push(sort);
                    console.log(`第${i + 1}行 Sort: ${sort}`);
                }
            }
        }

        console.log('现有ID列表:', existingIds);
        console.log('现有Sort列表:', existingSorts);

        const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 0;
        const maxSort = existingSorts.length > 0 ? Math.max(...existingSorts) : 0;
        const newId = maxId + 1;
        const newSort = maxSort + 1;

        console.log(`计算结果: 最大ID=${maxId}, 新ID=${newId}, 最大Sort=${maxSort}, 新Sort=${newSort}`);

        // 提取主题基础名称
        const baseName = extractThemeBaseName(themeName);
        console.log('提取的基础主题名称:', baseName);

        // 创建新行
        const newRow = new Array(headerRow.length).fill('');
        newRow[idColumnIndex] = newId.toString();
        newRow[notesColumnIndex] = baseName;
        newRow[nameIDColumnIndex] = multiLangId.toString();
        newRow[sortColumnIndex] = newSort.toString();
        newRow[isFilterColumnIndex] = '1';

        console.log('准备添加的新行数据:', newRow);
        console.log('新行将添加到第', infoData.length + 1, '行');

        infoData.push(newRow);

        console.log(`✅ 已添加新的AllObstacle记录: ID=${newId}, 基础名称=${baseName}, 多语言ID=${multiLangId}, Sort=${newSort}`);

        return {
            updated: true,
            data: infoData,
            message: `已添加AllObstacle记录: ID=${newId}, 主题=${baseName}, 多语言ID=${multiLangId}`
        };
    }

    /**
     * 保存AllObstacle文件
     * @param {Object} workbook - 更新后的工作簿
     */
    async function saveAllObstacleFileDirectly(workbook) {
        console.log('开始保存AllObstacle文件到原位置...');

        try {
            // 检查文件句柄权限
            const fileHandle = folderManager.allObstacleHandle;
            const permission = await fileHandle.queryPermission({ mode: 'readwrite' });
            console.log('AllObstacle文件当前权限:', permission);

            if (permission !== 'granted') {
                console.log('尝试重新请求AllObstacle文件权限...');
                const newPermission = await fileHandle.requestPermission({ mode: 'readwrite' });
                if (newPermission !== 'granted') {
                    throw new Error('无法获取AllObstacle文件写入权限');
                }
            }

            // 尝试保存为XLS格式（Java工具需要OLE2格式）
            let excelBuffer;
            let saveSuccess = false;

            // 优先尝试保存为XLS格式，因为Java工具需要OLE2格式
            const saveAttempts = [
                { bookType: 'xls', type: 'array', cellDates: false, description: 'XLS格式（无日期处理）' },
                { bookType: 'xls', type: 'array', cellDates: true, description: 'XLS格式（含日期处理）' }
            ];

            for (const attempt of saveAttempts) {
                try {
                    console.log(`尝试使用${attempt.description}保存...`);
                    excelBuffer = XLSX.write(workbook, attempt);
                    saveSuccess = true;
                    console.log(`✅ ${attempt.description}生成成功`);
                    break;
                } catch (writeError) {
                    console.warn(`❌ ${attempt.description}生成失败:`, writeError.message);
                    continue;
                }
            }

            if (!saveSuccess) {
                console.log('XLS格式保存失败，尝试重新构建工作簿...');
                try {
                    const cleanWorkbook = rebuildAllObstacleWorkbook(workbook);
                    // 重新构建后仍然尝试保存为XLS格式
                    excelBuffer = XLSX.write(cleanWorkbook, {
                        bookType: 'xls',
                        type: 'array',
                        cellDates: false
                    });
                    saveSuccess = true;
                    console.log('✅ 重新构建工作簿并保存为XLS格式成功');
                } catch (rebuildError) {
                    console.error('重新构建工作簿失败:', rebuildError);
                    // 最后尝试XLSX格式，但会警告用户
                    try {
                        excelBuffer = XLSX.write(workbook, {
                            bookType: 'xlsx',
                            type: 'array'
                        });
                        saveSuccess = true;
                        console.warn('⚠️ 已保存为XLSX格式，可能与Java工具不兼容');
                        App.Utils.showStatus('警告：AllObstacle文件已保存为XLSX格式，可能需要手动转换为XLS格式', 'warning', 8000);
                    } catch (xlsxError) {
                        throw new Error('所有保存格式都失败，文件包含不兼容的数据类型');
                    }
                }
            }

            // 创建可写流并写入数据
            const writable = await fileHandle.createWritable();
            await writable.write(excelBuffer);
            await writable.close();

            console.log('AllObstacle文件已成功保存');
            return true;

        } catch (error) {
            console.error('AllObstacle文件保存失败:', error);

            // 如果是XLSX库的类型错误，提供更友好的错误信息
            if (error.message.includes('TypedPropertyValue') || error.message.includes('unrecognized type')) {
                const friendlyError = new Error('AllObstacle.xls文件包含不兼容的数据格式，建议使用较新版本的Excel重新保存该文件');
                friendlyError.originalError = error;
                throw friendlyError;
            }

            throw error;
        }
    }

    /**
     * 重新构建AllObstacle工作簿以解决兼容性问题
     * @param {Object} originalWorkbook - 原始工作簿
     * @returns {Object} 重新构建的工作簿
     */
    function rebuildAllObstacleWorkbook(originalWorkbook) {
        console.log('开始重新构建AllObstacle工作簿...');

        try {
            const newWorkbook = XLSX.utils.book_new();

            // 遍历所有工作表
            for (const sheetName of originalWorkbook.SheetNames) {
                console.log(`重新构建工作表: ${sheetName}`);

                const originalSheet = originalWorkbook.Sheets[sheetName];

                // 将工作表转换为纯数据数组
                const sheetData = XLSX.utils.sheet_to_json(originalSheet, {
                    header: 1,
                    defval: '',
                    raw: true // 保持原始数据类型
                });

                // 清理数据，智能处理数据类型
                const cleanData = sheetData.map((row, rowIndex) => {
                    if (!Array.isArray(row)) return [];
                    return row.map((cell, colIndex) => {
                        if (cell === null || cell === undefined) return '';

                        // 如果是复杂对象，转换为字符串
                        if (typeof cell === 'object' && cell.constructor !== Date) {
                            return String(cell);
                        }

                        // 如果是日期对象，转换为字符串
                        if (cell instanceof Date) {
                            return cell.toISOString().split('T')[0]; // YYYY-MM-DD格式
                        }

                        // 对于数字，保持数字类型但确保是有效数字
                        if (typeof cell === 'number') {
                            if (isNaN(cell) || !isFinite(cell)) {
                                return '';
                            }
                            return cell;
                        }

                        // 对于字符串，检查是否是数字字符串
                        if (typeof cell === 'string') {
                            const trimmed = cell.trim();
                            if (trimmed === '') return '';

                            // 尝试转换为数字（仅对纯数字字符串）
                            const num = Number(trimmed);
                            if (!isNaN(num) && isFinite(num) && /^\d+\.?\d*$/.test(trimmed)) {
                                return num;
                            }
                            return trimmed;
                        }

                        return cell;
                    });
                });

                // 创建新的工作表
                const newSheet = XLSX.utils.aoa_to_sheet(cleanData);

                // 添加到新工作簿
                XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);

                console.log(`✅ 工作表 ${sheetName} 重新构建完成`);
            }

            console.log('AllObstacle工作簿重新构建完成');
            return newWorkbook;

        } catch (error) {
            console.error('重新构建AllObstacle工作簿失败:', error);
            throw error;
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
        rscOriginalSheetsData = null; // 🔧 清理原始数据备份
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
            rscOriginalSheetsData = {}; // 🔧 同时保存原始数据的深拷贝
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const sheetData = XLSX.utils.sheet_to_json(worksheet, {
                    header: 1,
                    defval: '',
                    raw: false
                });
                rscAllSheetsData[sheetName] = sheetData;

                // 🔧 深拷贝原始数据（用于后续重置非目标工作表）
                rscOriginalSheetsData[sheetName] = JSON.parse(JSON.stringify(sheetData));
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

            // 同步目标工作表（严格限制：仅限Light、ColorInfo、FloodLight、VolumetricFog）
            // 重要约束：不同步其他工作表，保持零影响原则
            // 🔧 修复：根据Status工作表状态获取需要处理的工作表列表
            const targetSheets = getActiveSheetsByStatus(false); // syncMemoryDataState中传递false
            console.log('🎯 根据Status状态确定的目标工作表（syncMemoryDataState）:', targetSheets);

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
        getCurrentTimeString: getCurrentTimeString,

        // 测试功能（用于验证修改）
        detectMappingMode: detectMappingMode,
        parseStatusSheet: parseStatusSheet,
        findLightValueDirect: findLightValueDirect,
        findLightValueFromSourceLight: findLightValueFromSourceLight,
        findLightValueFromRSCThemeLight: findLightValueFromRSCThemeLight,
        loadExistingLightConfig: loadExistingLightConfig,
        findColorInfoValueDirect: findColorInfoValueDirect,
        findColorInfoValueFromSourceColorInfo: findColorInfoValueFromSourceColorInfo,
        findColorInfoValueFromRSCThemeColorInfo: findColorInfoValueFromRSCThemeColorInfo,
        loadExistingColorInfoConfig: loadExistingColorInfoConfig,
        findVolumetricFogValueDirect: findVolumetricFogValueDirect,
        findVolumetricFogValueFromSourceVolumetricFog: findVolumetricFogValueFromSourceVolumetricFog,
        findVolumetricFogValueFromRSCThemeVolumetricFog: findVolumetricFogValueFromRSCThemeVolumetricFog,
        loadExistingVolumetricFogConfig: loadExistingVolumetricFogConfig
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

            // 设置AllObstacle文件信息
            if (result.allObstacleFound && result.files.allObstacle.hasPermission) {
                try {
                    const allObstacleFileData = await folderManager.loadThemeFileData('allObstacle');
                    await setAllObstacleDataFromFolder(allObstacleFileData);
                    console.log('AllObstacle文件加载成功');
                } catch (error) {
                    console.error('AllObstacle文件加载失败:', error);
                    App.Utils.showStatus('AllObstacle文件加载失败，AllObstacle功能将不可用', 'warning', 5000);
                }
            } else if (result.allObstacleFound) {
                console.warn('AllObstacle.xls文件找到但权限获取失败');
                App.Utils.showStatus('AllObstacle文件权限获取失败，AllObstacle功能将不可用', 'warning', 5000);
            } else {
                console.log('未找到AllObstacle.xls文件，AllObstacle功能将不可用');
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
            rscOriginalSheetsData = {}; // 🔧 同时保存原始数据的深拷贝
            fileData.workbook.SheetNames.forEach(sheetName => {
                const sheet = fileData.workbook.Sheets[sheetName];
                const sheetData = XLSX.utils.sheet_to_json(sheet, {
                    header: 1,
                    defval: '',
                    raw: false
                });
                rscAllSheetsData[sheetName] = sheetData;

                // 🔧 深拷贝原始数据（用于后续重置非目标工作表）
                rscOriginalSheetsData[sheetName] = JSON.parse(JSON.stringify(sheetData));
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
     * 从文件夹模式设置AllObstacle数据
     */
    async function setAllObstacleDataFromFolder(allObstacleFileData) {
        try {
            console.log('设置AllObstacle数据...');

            allObstacleData = {
                workbook: allObstacleFileData.workbook,
                fileHandle: allObstacleFileData.fileHandle,
                fileName: allObstacleFileData.fileName,
                fileSize: allObstacleFileData.fileSize,
                lastModified: allObstacleFileData.lastModified
            };

            console.log('AllObstacle数据设置成功:', {
                fileName: allObstacleData.fileName,
                fileSize: allObstacleData.fileSize,
                sheetCount: allObstacleData.workbook.SheetNames.length
            });

            App.Utils.showStatus('AllObstacle.xls文件已加载，支持全新系列主题的AllObstacle处理', 'info', 3000);

        } catch (error) {
            console.error('AllObstacle数据设置失败:', error);
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
     * @param {string} themeName - 主题名称
     * @param {boolean} isNewTheme - 是否为新建主题（默认false）
     */
    function loadExistingLightConfig(themeName, isNewTheme = false) {
        console.log('加载Light配置:', themeName);
        console.log('是否新建主题:', isNewTheme);
        console.log('当前映射模式:', currentMappingMode);

        // 检查是否为直接映射模式
        const isDirectMode = currentMappingMode === 'direct';

        if (isDirectMode) {
            console.log('直接映射模式：优先从源数据Light工作表读取配置显示');

            // Light字段映射
            const lightFieldMapping = {
                'Max': 'lightMax',
                'Dark': 'lightDark',
                'Min': 'lightMin',
                'SpecularLevel': 'lightSpecularLevel',
                'Gloss': 'lightGloss',
                'SpecularColor': 'lightSpecularColor'
            };

            let hasSourceData = false;

            // 尝试从源数据Light工作表读取每个字段
            Object.entries(lightFieldMapping).forEach(([lightColumn, inputId]) => {
                // 🔧 使用条件读取逻辑获取Light字段值，传递正确的isNewTheme参数
                const directValue = findLightValueDirect(lightColumn, isNewTheme, themeName);

                const input = document.getElementById(inputId);
                if (input) {
                    if (directValue !== null && directValue !== undefined && directValue !== '') {
                        input.value = directValue;
                        hasSourceData = true;
                        console.log(`✅ 直接映射模式从条件读取获取Light配置: ${lightColumn} = ${directValue}`);

                        // 更新颜色预览
                        if (inputId === 'lightSpecularColor' && window.App && window.App.ColorPicker) {
                            window.App.ColorPicker.updateColorPreview(inputId, directValue);
                        }
                    } else {
                        // 如果条件读取失败，使用默认值
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

        // 非直接映射模式或直接映射模式回退：从RSC_Theme读取
        console.log('从RSC_Theme Light工作表读取配置');

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
     * 获取最后一个主题的FloodLight配置（用于非直接映射模式的默认值）
     * @returns {Object} FloodLight配置对象
     */
    function getLastThemeFloodLightConfig() {
        console.log('=== 获取最后一个主题的FloodLight配置 ===');

        try {
            if (!rscAllSheetsData || !rscAllSheetsData['FloodLight']) {
                console.log('RSC_Theme FloodLight数据未加载');
                return {};
            }

            const floodLightData = rscAllSheetsData['FloodLight'];
            if (!floodLightData || floodLightData.length < 2) {
                console.log('FloodLight工作表数据不足');
                return {};
            }

            const headerRow = floodLightData[0];
            const lastDataRow = floodLightData[floodLightData.length - 1];

            const config = {};
            headerRow.forEach((columnName, index) => {
                if (columnName && columnName !== 'id' && columnName !== 'notes') {
                    config[columnName] = lastDataRow[index] || '';
                }
            });

            console.log('最后一个主题的FloodLight配置:', config);
            return config;

        } catch (error) {
            console.error('获取最后一个主题的FloodLight配置时出错:', error);
            return {};
        }
    }

    /**
     * 加载现有主题的FloodLight配置
     * @param {string} themeName - 主题名称
     * @param {boolean} isNewTheme - 是否为新建主题（默认false）
     */
    function loadExistingFloodLightConfig(themeName, isNewTheme = false) {
        console.log('加载FloodLight配置:', themeName);
        console.log('是否新建主题:', isNewTheme);

        // 检测映射模式
        const isDirectMode = currentMappingMode === 'direct';
        console.log(`当前映射模式: ${currentMappingMode}, 是否直接映射: ${isDirectMode}`);

        // FloodLight字段映射
        const floodLightFieldMapping = {
            'Color': 'floodlightColor',
            'TippingPoint': 'floodlightTippingPoint',
            'Strength': 'floodlightStrength',
            'IsOn': 'floodlightIsOn',
            'JumpActiveIsLightOn': 'floodlightJumpActiveIsLightOn',
            'LightStrength': 'floodlightLightStrength'
        };

        if (isDirectMode && sourceData && sourceData.workbook) {
            // 直接映射模式：优先使用条件读取逻辑
            console.log('直接映射模式：使用条件读取逻辑加载FloodLight配置');

            Object.entries(floodLightFieldMapping).forEach(([columnName, fieldId]) => {
                const input = document.getElementById(fieldId);
                if (!input) {
                    console.warn(`未找到UI元素: ${fieldId}`);
                    return;
                }

                // 🔧 使用条件读取逻辑获取字段值，传递正确的isNewTheme参数
                const value = findFloodLightValueDirect(columnName, isNewTheme, themeName);

                if (value !== null && value !== undefined && value !== '') {
                    if (fieldId === 'floodlightTippingPoint' || fieldId === 'floodlightStrength') {
                        // 将存储的整数值转换为小数显示（除以10）
                        const numValue = parseInt(value) || 0;
                        input.value = (numValue / 10).toFixed(1);
                        console.log(`✅ 直接映射加载: ${columnName} = ${input.value} (原始值: ${value})`);
                    } else if (fieldId === 'floodlightIsOn' || fieldId === 'floodlightJumpActiveIsLightOn') {
                        // 处理checkbox
                        input.checked = value === 1 || value === '1' || value === true;
                        console.log(`✅ 直接映射加载: ${columnName} = ${input.checked}`);
                    } else {
                        // 颜色值
                        input.value = value.toString();
                        console.log(`✅ 直接映射加载: ${columnName} = ${value}`);

                        // 更新颜色预览
                        if (fieldId === 'floodlightColor') {
                            updateFloodLightColorPreview();
                        }
                    }
                } else {
                    console.log(`⚠️ 直接映射未找到字段值: ${columnName}，保持UI默认值`);
                }
            });

            console.log('✅ 直接映射模式FloodLight配置加载完成');
            return;
        }

        // 非直接映射模式：使用原有逻辑从RSC_Theme加载
        if (!rscAllSheetsData || !rscAllSheetsData['FloodLight']) {
            console.log('RSC_Theme FloodLight数据未加载，使用默认值');
            resetFloodLightConfigToDefaults();
            return;
        }

        // 在RSC_Theme的FloodLight sheet中查找主题
        const floodLightData = rscAllSheetsData['FloodLight'];
        const floodLightHeaderRow = floodLightData[0];
        const floodLightNotesColumnIndex = floodLightHeaderRow.findIndex(col => col === 'notes');

        if (floodLightNotesColumnIndex === -1) {
            console.log('RSC_Theme FloodLight sheet没有notes列，使用默认值');
            resetFloodLightConfigToDefaults();
            return;
        }

        // 查找主题在FloodLight中的行号
        const floodLightThemeRowIndex = floodLightData.findIndex((row, index) =>
            index > 0 && row[floodLightNotesColumnIndex] === themeName
        );

        if (floodLightThemeRowIndex === -1) {
            console.log(`在RSC_Theme FloodLight sheet中未找到主题 "${themeName}"，使用默认值`);
            resetFloodLightConfigToDefaults();
            return;
        }

        console.log(`在RSC_Theme FloodLight sheet中找到主题 "${themeName}"，行索引: ${floodLightThemeRowIndex}`);

        // 加载FloodLight配置值
        const floodLightRow = floodLightData[floodLightThemeRowIndex];

        Object.entries(floodLightFieldMapping).forEach(([columnName, fieldId]) => {
            const columnIndex = floodLightHeaderRow.findIndex(col => col === columnName);
            if (columnIndex !== -1) {
                const value = floodLightRow[columnIndex];
                const input = document.getElementById(fieldId);

                if (input) {
                    if (fieldId === 'floodlightTippingPoint' || fieldId === 'floodlightStrength') {
                        // 将存储的整数值转换为小数显示（除以10）
                        const numValue = parseInt(value) || 0;
                        input.value = (numValue / 10).toFixed(1);
                    } else if (fieldId === 'floodlightIsOn' || fieldId === 'floodlightJumpActiveIsLightOn') {
                        // 处理checkbox
                        input.checked = value === 1 || value === '1' || value === true;
                    } else {
                        // 颜色值
                        input.value = (value !== undefined && value !== null && value !== '') ? value.toString() : 'FFFFFF';

                        // 更新颜色预览
                        if (fieldId === 'floodlightColor') {
                            updateFloodLightColorPreview();
                        }
                    }
                }
            } else {
                console.warn(`FloodLight sheet中找不到列: ${columnName}`);
            }
        });

        console.log('FloodLight配置加载完成');
    }

    /**
     * 加载现有主题的VolumetricFog配置
     * @param {string} themeName - 主题名称
     * @param {boolean} isNewTheme - 是否为新建主题（默认false）
     */
    function loadExistingVolumetricFogConfig(themeName, isNewTheme = false) {
        console.log('加载VolumetricFog配置:', themeName);
        console.log('是否新建主题:', isNewTheme);
        console.log('当前映射模式:', currentMappingMode);

        // 检查是否为直接映射模式
        const isDirectMode = currentMappingMode === 'direct';

        if (isDirectMode) {
            // 直接映射模式：使用条件读取逻辑显示源数据VolumetricFog配置
            console.log('直接映射模式：尝试从源数据VolumetricFog工作表加载配置');

            // 检查Status工作表中VolumetricFog状态
            let volumetricFogStatusFromStatus = 0;
            if (sourceData && sourceData.workbook) {
                const statusInfo = parseStatusSheet(sourceData);
                volumetricFogStatusFromStatus = statusInfo.volumetricFogStatus;
                console.log(`Status工作表VolumetricFog状态: ${volumetricFogStatusFromStatus}`);
            }

            // VolumetricFog字段映射
            const volumetricFogFieldMapping = {
                'Color': 'volumetricfogColor',
                'X': 'volumetricfogX',
                'Y': 'volumetricfogY',
                'Z': 'volumetricfogZ',
                'Density': 'volumetricfogDensity',
                'Rotate': 'volumetricfogRotate',
                'IsOn': 'volumetricfogIsOn'
            };

            let hasSourceData = false;

            // 尝试从源数据VolumetricFog工作表读取每个字段
            Object.entries(volumetricFogFieldMapping).forEach(([columnName, inputId]) => {
                const input = document.getElementById(inputId);
                if (!input) return;

                // 特殊处理IsOn字段：如果Status工作表中VolumetricFog状态为1，则自动勾选
                if (inputId === 'volumetricfogIsOn') {
                    if (volumetricFogStatusFromStatus === 1) {
                        input.checked = true;
                        hasSourceData = true;
                        console.log(`✅ Status工作表VolumetricFog状态为1，自动勾选IsOn: ${columnName} = true`);
                        return; // 跳过后续的条件读取逻辑
                    }
                }

                // 🔧 使用条件读取逻辑获取VolumetricFog字段值，传递正确的isNewTheme参数
                const directValue = findVolumetricFogValueDirect(columnName, isNewTheme, themeName);

                if (directValue !== null && directValue !== undefined && directValue !== '') {
                    hasSourceData = true;

                    if (inputId === 'volumetricfogIsOn') {
                        // 处理checkbox
                        input.checked = directValue === 1 || directValue === '1' || directValue === true;
                    } else if (inputId === 'volumetricfogDensity') {
                        // Density字段需要÷10显示
                        const displayValue = (parseFloat(directValue) / 10).toFixed(1);
                        input.value = displayValue;
                    } else {
                        // 其他字段
                        input.value = directValue.toString();

                        // 更新颜色预览
                        if (inputId === 'volumetricfogColor') {
                            updateVolumetricFogColorPreview();
                        }
                    }

                    console.log(`✅ 直接映射加载VolumetricFog字段: ${columnName} = ${directValue}`);
                } else {
                    console.log(`直接映射未找到VolumetricFog字段: ${columnName}，将使用RSC_Theme数据`);
                }
            });

            if (hasSourceData) {
                console.log('✅ 直接映射模式：成功从源数据VolumetricFog工作表加载配置');
                return;
            } else {
                console.log('直接映射模式：源数据VolumetricFog工作表无可用数据，回退到RSC_Theme数据');
            }
        }

        // 非直接映射模式或直接映射模式回退：从RSC_Theme加载
        if (!rscAllSheetsData || !rscAllSheetsData['VolumetricFog']) {
            console.log('RSC_Theme VolumetricFog数据未加载，使用默认值');
            resetVolumetricFogConfigToDefaults();
            return;
        }

        // 在RSC_Theme的VolumetricFog sheet中查找主题
        const volumetricFogData = rscAllSheetsData['VolumetricFog'];
        const volumetricFogHeaderRow = volumetricFogData[0];
        const volumetricFogNotesColumnIndex = volumetricFogHeaderRow.findIndex(col => col === 'notes');

        if (volumetricFogNotesColumnIndex === -1) {
            console.log('RSC_Theme VolumetricFog sheet没有notes列，使用默认值');
            resetVolumetricFogConfigToDefaults();
            return;
        }

        // 查找主题行
        let themeRowIndex = -1;
        for (let i = 1; i < volumetricFogData.length; i++) {
            if (volumetricFogData[i][volumetricFogNotesColumnIndex] === themeName) {
                themeRowIndex = i;
                break;
            }
        }

        if (themeRowIndex === -1) {
            console.log(`在VolumetricFog sheet中未找到主题"${themeName}"，使用默认值`);
            resetVolumetricFogConfigToDefaults();
            return;
        }

        const themeRow = volumetricFogData[themeRowIndex];
        console.log(`找到VolumetricFog主题"${themeName}"，行索引: ${themeRowIndex}`, themeRow);

        // VolumetricFog字段映射
        const volumetricFogFieldMapping = {
            'Color': 'volumetricfogColor',
            'X': 'volumetricfogX',
            'Y': 'volumetricfogY',
            'Z': 'volumetricfogZ',
            'Density': 'volumetricfogDensity',
            'Rotate': 'volumetricfogRotate',
            'IsOn': 'volumetricfogIsOn'
        };

        // 加载配置到界面
        Object.entries(volumetricFogFieldMapping).forEach(([columnName, fieldId]) => {
            const columnIndex = volumetricFogHeaderRow.findIndex(col => col === columnName);
            if (columnIndex !== -1) {
                const value = themeRow[columnIndex];
                const input = document.getElementById(fieldId);

                if (input) {
                    if (fieldId === 'volumetricfogIsOn') {
                        // 处理checkbox
                        input.checked = value === 1 || value === '1' || value === true;
                    } else if (fieldId === 'volumetricfogDensity') {
                        // Density字段需要÷10显示
                        const displayValue = (parseFloat(value) / 10).toFixed(1);
                        input.value = displayValue;
                    } else {
                        // 其他字段
                        input.value = (value !== undefined && value !== null && value !== '') ? value.toString() : '';

                        // 更新颜色预览
                        if (fieldId === 'volumetricfogColor') {
                            updateVolumetricFogColorPreview();
                        }
                    }
                }
            } else {
                console.warn(`VolumetricFog sheet中找不到列: ${columnName}`);
            }
        });

        console.log('VolumetricFog配置加载完成');
    }

    /**
     * 加载现有主题的ColorInfo配置
     * @param {string} themeName - 主题名称
     * @param {boolean} isNewTheme - 是否为新建主题（默认false）
     */
    function loadExistingColorInfoConfig(themeName, isNewTheme = false) {
        console.log('加载ColorInfo配置:', themeName);
        console.log('是否新建主题:', isNewTheme);
        console.log('当前映射模式:', currentMappingMode);

        // 检查是否为直接映射模式
        const isDirectMode = currentMappingMode === 'direct';

        if (isDirectMode) {
            console.log('直接映射模式：优先从源数据ColorInfo工作表读取配置显示');

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

            let hasSourceData = false;

            // 尝试从源数据ColorInfo工作表读取每个字段
            Object.entries(colorInfoFieldMapping).forEach(([columnName, inputId]) => {
                // 🔧 使用条件读取逻辑获取ColorInfo字段值，传递正确的isNewTheme参数
                const directValue = findColorInfoValueDirect(columnName, isNewTheme, themeName);

                const input = document.getElementById(inputId);
                if (input) {
                    if (directValue !== null && directValue !== undefined && directValue !== '') {
                        input.value = directValue;
                        hasSourceData = true;
                        console.log(`✅ 直接映射模式从条件读取获取ColorInfo配置: ${columnName} = ${directValue}`);
                    } else {
                        // 如果条件读取失败，使用默认值
                        const colorInfoDefaults = getLastThemeColorInfoConfig();
                        const defaultValue = colorInfoDefaults[inputId] || '0';
                        input.value = defaultValue;
                        console.log(`⚠️ 直接映射模式ColorInfo字段条件读取失败，使用默认值: ${columnName} = ${defaultValue}`);
                    }
                }
            });

            if (hasSourceData) {
                console.log('✅ 直接映射模式：成功从源数据加载ColorInfo配置');

                // 更新颜色预览
                updateRgbColorPreview('PickupDiffR');
                updateRgbColorPreview('PickupReflR');
                updateRgbColorPreview('BallSpecR');
                updateRgbColorPreview('ForegroundFogR');

                return;
            } else {
                console.log('⚠️ 直接映射模式：未能从源数据获取ColorInfo配置，回退到RSC_Theme读取');
            }
        }

        // 非直接映射模式或直接映射模式回退：从RSC_Theme读取
        console.log('从RSC_Theme ColorInfo工作表读取配置');

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
