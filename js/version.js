/**
 * 版本管理模块
 * 文件说明：管理应用版本信息和版本历史
 * 创建时间：2025-01-09
 * 依赖：无
 */

// 创建全局命名空间
window.App = window.App || {};

/**
 * 版本管理模块
 * 提供版本信息显示和管理功能
 */
window.App.Version = (function() {
    'use strict';

    // 当前版本信息
    const VERSION_INFO = {
        current: '1.4.8',
        releaseDate: '2025-09-18',
        buildNumber: '20250918001'
    };

    // 版本历史记录
    const VERSION_HISTORY = [
        {
            version: '1.4.8',
            date: '2025-09-18',
            changes: [
                '严格限制RSC_Theme工作表的修改范围，不要改到无关表',
            ],
            type: 'patch'
        },
        {
            version: '1.4.7',
            date: '2025-09-16',
            changes: [
                '修正UGCTheme表Level_show_id字段填值逻辑，根据主题类型采用不同策略',
                '全新主题系列：Level_show_id设置为"新增主题ID - 1"',
                '现有主题系列：Level_show_id继续使用"上一行数据值 + 1"的递增逻辑',
                '基于智能多语言配置系统实现主题类型自动识别',
                '添加详细的填值过程日志输出，包含计算来源和处理步骤',
                '确保与多语言配置功能完全兼容，不影响现有功能',
                '提升数据处理准确性，避免全新主题错误延续递增逻辑',
                '完善错误处理机制，处理列不存在和数据异常情况'
            ],
            type: 'patch'
        },
        {
            version: '1.4.6',
            date: '2025-09-12',
            changes: [
                '新增源数据文件选择结果显示功能，提供文件名、大小、时间和数据行数的详细信息',
                '优化UI界面，移除冗余的文件状态栏和右侧信息区域，界面更加简洁',
                '修改最终操作指引弹窗，使用固定路径显示，提供更明确的操作指引',
                '完善文件选择状态管理，在错误情况下自动隐藏结果显示',
                '集成文件选择结果到重置功能，确保界面状态一致性',
                '使用绿色主题的成功状态显示，与项目整体UI风格保持一致',
                '添加实时文件信息更新，包括格式化的文件大小和精确的选择时间',
                '提供完整的错误处理机制，确保用户体验的流畅性'
            ],
            type: 'minor'
        },
        {
            version: '1.4.5',
            date: '2025-09-12',
            changes: [
                '更新图案'
            ],
            type: 'minor'
        },
        {
            version: '1.4.4',
            date: '2025-09-10',
            changes: [
                '新增BallSpec钻石高光颜色配置功能，支持钻石高光颜色的RGB设置',
                '在ColorInfo配置面板中添加BallSpecR、BallSpecG、BallSpecB三个新字段',
                '完整实现BallSpec字段的UI界面、数据收集、数据应用、验证和默认值处理',
                '集成BallSpec字段到getColorInfoConfigData()数据收集函数',
                '集成BallSpec字段到applyColorInfoConfigToRow()数据应用函数',
                '集成BallSpec字段到initColorInfoValidation()验证函数',
                '集成BallSpec字段到getLastThemeColorInfoConfig()默认值函数',
                '集成BallSpec字段到updateRgbColorPreview()颜色预览函数',
                '修复loadExistingColorInfoConfig()函数缺少BallSpec字段映射的问题',
                '确保新建和更新主题时BallSpec配置都能正确保存到RSC_Theme文件',
                '提供完整的BallSpec颜色选择器，支持可视化颜色选择和RGB值自动拆分',
                '添加BallSpec字段的实时验证，确保RGB值在0-255范围内',
                '完善ColorInfo配置功能，现在支持钻石颜色、反光颜色、高光颜色和雾效设置'
            ],
            type: 'minor'
        },
        {
            version: '1.4.3',
            date: '2025-09-10',
            changes: [
                '修复Light和ColorInfo数据保存的关键问题，确保配置数据正确写入Excel文件',
                '解决新建主题时Light和ColorInfo工作表不创建新行的问题',
                '修复更新现有主题时Light和ColorInfo配置被忽略的问题',
                '修复generateUpdatedWorkbook函数只处理主工作表的缺陷，现在正确处理所有工作表',
                '修复processRSCAdditionalSheets函数在新建主题时没有更新rscAllSheetsData的问题',
                '新增updateExistingThemeAdditionalSheets函数，专门处理现有主题的Light和ColorInfo更新',
                '完善数据流程，确保无论新建还是更新主题，Light和ColorInfo配置都能正确保存',
                '移除过时的功能限制说明，现在所有配置功能都完全可用',
                '优化数据同步逻辑，确保workbook.Sheets和rscAllSheetsData保持一致',
                '添加详细的调试日志，便于问题诊断和数据流程追踪'
            ],
            type: 'major'
        },
        {
            version: '1.4.2',
            date: '2025-09-10',
            changes: [
                '初步最终集成工具完成',
            ],
            type: 'minor'
        },
        {
            version: '1.4.1',
            date: '2025-01-10',
            changes: [
                '新增RSC_Theme ColorInfo配置功能，支持钻石颜色和远景雾颜色设置',
                '实现RGB颜色字段的可视化配置，支持钻石颜色、反光颜色、远景雾颜色',
                '添加雾效距离参数配置，支持FogStart和FogEnd的范围验证',
                '为跳板边框图案ID添加可视化选择器，与地板边框保持一致的交互体验',
                '修复玻璃透明度默认值问题，确保新建主题时默认值为50',
                '优化Light和ColorInfo配置的默认值逻辑，新建主题时使用表中最后一个主题的数据',
                '移除跳板边框图案的0/1限制注释，支持任意有效的边框图案ID选择',
                '完善RGB颜色选择器，支持颜色自动拆分为R、G、B三个数值字段',
                '增强配置面板的用户体验，提供实时颜色预览和验证反馈'
            ],
            type: 'minor'
        },
        {
            version: '1.4.0',
            date: '2025-01-09',
            changes: [
                '新增RSC_Theme Light配置功能，支持完整的光照参数设置',
                '实现明度偏移、高光等级、光泽度和高光颜色的可视化配置',
                '添加专业的16进制颜色选择器，支持渐变色板和预设颜色',
                '集成实时数值验证系统，自动检查范围限制并提供错误提示',
                '修复UGC图案选择器的事件绑定问题，确保按钮响应正常',
                '优化Light配置面板UI设计，提供清晰的字段说明和使用提示',
                '添加Light配置的数据加载和保存功能，支持现有主题值显示'
            ],
            type: 'major'
        },
        {
            version: '1.3.1',
            date: '2025-09-09',
            changes: [
                '更换主题Unity工具链接',
                '更新导航栏UI'
            ],
            type: 'minor'
        },
        {
            version: '1.3.0',
            date: '2025-09-09',
            changes: [
                '增加导航站',
            ],
            type: 'minor'
        },
        {
            version: '1.2.9',
            date: '2025-09-09',
            changes: [
                '更换静态部署工具为GitLab',
            ],
            type: 'minor'
        },
        {
            version: '1.2.8',
            date: '2025-09-08',
            changes: [
                '手动改表的操作写到工具的自动化流程中',
            ],
            type: 'minor'
        },
        {
            version: '1.2.7',
            date: '2025-09-08',
            changes: [
                '实现多语言配置智能检测逻辑，支持主题名称相似性检测',
                '添加同系列主题自动复用多语言配置功能',
                '优化多语言配置面板显示逻辑，仅在全新主题系列时显示',
                '新增主题类型指示器，提供清晰的状态反馈',
                '支持中英文主题名称的智能基础名称提取',
                '完善UGCTheme处理中的多语言ID复用机制'
            ],
            type: 'minor'
        },
        {
            version: '1.2.6',
            date: '2025-09-08',
            changes: [
                '实现UGCTheme文件多语言ID自动填充功能',
                '在processUGCTheme函数中添加LevelName列的智能处理',
                '支持多种LevelName列名格式的自动识别',
                '用户输入有效时使用多语言ID，无效时使用上一行数据',
                '添加详细的日志输出和错误处理机制',
                '修复多语言配置UI样式问题，提高文字可见性'
            ],
            type: 'minor'
        },
        {
            version: '1.2.5',
            date: '2025-09-08',
            changes: [
                '新增多语言配置UI功能，仅在创建新主题时显示',
                '添加主题显示名称输入框和多语言ID输入框',
                '集成在线多语言表链接，提供清晰的操作指引',
                '实现多语言配置验证逻辑，不影响现有核心功能',
                '优化用户界面，保持与现有风格一致'
            ],
            type: 'minor'
        },
        {
            version: '1.2.4',
            date: '2025-09-08',
            changes: [
                'UGCTheme表的Level_show_id与Level_show_bg_ID列不再需要手动操作,而是放到当前工具内的自动化流程里面了',
            ],
            type: 'minor'
        },
        {
            version: '1.2.3',
            date: '2025-09-05',
            changes: [
                '重构映射逻辑：优先检查颜色代码，再验证RC通道有效性',
                '改进映射处理：支持更灵活的颜色代码到RC通道的映射',
                '优化跳过逻辑：只有颜色代码为空时才跳过，RC通道无效时使用默认值',
                '增强日志输出：提供更详细的映射处理统计信息',
                '提升映射覆盖率：处理更多映射表中的有效颜色代码'
            ],
            type: 'minor'
        },
        {
            version: '1.2.1',
            date: '2025-09-05',
            changes: [
                '文字修改'
            ],
            type: 'minor'
        },
        {
            version: '1.2.0',
            date: '2025-09-05',
            changes: [
                '修复主题提取逻辑，正确从第6行开始读取主题数据',
                '优化更新模式下UGCTheme文件处理逻辑',
                '改进错误提示，明确指出缺少的具体文件',
                '完善主题数据处理流程和用户反馈',
                '添加功能限制说明：部分数据如雾色花纹使用上一个主题的数据'
            ],
            type: 'minor'
        },
        {
            version: '1.1.1',
            date: '2025-09-05',
            changes: [
                '修复Excel文件格式兼容性问题',
                '统一使用.xls格式以兼容Unity工具',
                '解决Apache POI HSSF库读取错误',
                '添加版本显示功能',
                '优化文件保存和下载流程'
            ],
            type: 'major'
        },
        {
            version: '1.0.0',
            date: '2025-01-08',
            changes: [
                '初始版本发布',
                '实现颜色主题数据处理功能',
                '支持RSC_Theme和UGCTheme文件管理',
                '添加File System Access API支持',
                '实现数据预览和Sheet选择功能',
                '添加响应式界面设计',
                '支持拖拽文件上传',
                '实现数据同步和映射功能'
            ],
            type: 'major'
        }
    ];

    /**
     * 初始化版本模块
     */
    function init() {
        displayVersion();
        console.log(`颜色主题管理工具 v${VERSION_INFO.current} (${VERSION_INFO.releaseDate})`);
    }

    /**
     * 显示版本信息
     */
    function displayVersion() {
        const versionElement = document.getElementById('versionNumber');
        if (versionElement) {
            versionElement.textContent = `v${VERSION_INFO.current}`;
            versionElement.title = `版本 ${VERSION_INFO.current}\n发布日期: ${VERSION_INFO.releaseDate}\n构建号: ${VERSION_INFO.buildNumber}`;
        }
    }

    /**
     * 获取当前版本信息
     * @returns {Object} 版本信息对象
     */
    function getCurrentVersion() {
        return { ...VERSION_INFO };
    }

    /**
     * 获取版本历史
     * @returns {Array} 版本历史数组
     */
    function getVersionHistory() {
        return [...VERSION_HISTORY];
    }

    /**
     * 获取最新的更新内容
     * @returns {Object} 最新版本的更新内容
     */
    function getLatestChanges() {
        return VERSION_HISTORY[0] || null;
    }

    /**
     * 检查是否为新版本
     * @param {string} lastKnownVersion - 上次已知版本
     * @returns {boolean} 是否为新版本
     */
    function isNewVersion(lastKnownVersion) {
        if (!lastKnownVersion) return true;
        
        const current = VERSION_INFO.current.split('.').map(Number);
        const last = lastKnownVersion.split('.').map(Number);
        
        for (let i = 0; i < Math.max(current.length, last.length); i++) {
            const currentPart = current[i] || 0;
            const lastPart = last[i] || 0;
            
            if (currentPart > lastPart) return true;
            if (currentPart < lastPart) return false;
        }
        
        return false;
    }

    /**
     * 显示版本更新通知
     */
    function showUpdateNotification() {
        const lastVersion = localStorage.getItem('colorTool_lastVersion');
        
        if (isNewVersion(lastVersion)) {
            const changes = getLatestChanges();
            if (changes && App.Utils) {
                const changesList = changes.changes.slice(0, 3).join('、');
                App.Utils.showStatus(
                    `🎉 已更新到 v${VERSION_INFO.current}！主要更新：${changesList}`, 
                    'success', 
                    5000
                );
            }
            
            // 保存当前版本
            localStorage.setItem('colorTool_lastVersion', VERSION_INFO.current);
        }
    }

    // 暴露公共接口
    return {
        init: init,
        getCurrentVersion: getCurrentVersion,
        getVersionHistory: getVersionHistory,
        getLatestChanges: getLatestChanges,
        isNewVersion: isNewVersion,
        showUpdateNotification: showUpdateNotification
    };

})();

// 模块加载完成日志
console.log('Version模块已加载 - 版本管理功能已就绪');
