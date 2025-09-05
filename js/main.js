/**
 * 主应用模块
 * 文件说明：应用程序的主入口，负责初始化所有模块和协调模块间的交互
 * 创建时间：2025-01-09
 * 依赖：App.Utils, App.FileHandler, App.DataParser, App.TableRenderer, App.ExportUtils
 */

// 确保全局命名空间存在
window.App = window.App || {};

/**
 * 主应用模块
 * 负责应用程序的初始化和模块协调
 */
window.App.Main = (function() {
    'use strict';

    // 应用状态
    let isInitialized = false;
    let initializationAttempts = 0;
    const maxInitAttempts = 3;

    /**
     * 应用程序初始化
     * 按顺序初始化所有模块并设置全局事件监听器
     */
    function init() {
        if (isInitialized) {
            console.warn('应用程序已经初始化');
            return;
        }

        initializationAttempts++;
        console.log(`开始初始化应用程序 (第${initializationAttempts}次尝试)`);

        try {
            // 检查浏览器兼容性
            if (!checkCompatibility()) {
                return;
            }

            // 按依赖顺序初始化模块
            initializeModules();

            // 设置全局事件监听器
            setupGlobalEventListeners();

            // 设置全局函数（供HTML调用）
            setupGlobalFunctions();

            isInitialized = true;
            console.log('应用程序初始化完成');
            
            // 显示欢迎消息
            if (App.Utils) {
                App.Utils.showStatus('颜色主题管理工具已加载完成，请选择源数据文件和Unity项目', 'info');
            }

        } catch (error) {
            console.error('应用程序初始化失败:', error);
            
            if (initializationAttempts < maxInitAttempts) {
                console.log(`将在2秒后重试初始化...`);
                setTimeout(init, 2000);
            } else {
                console.error('应用程序初始化失败，已达到最大重试次数');
                showFatalError('应用程序初始化失败，请刷新页面重试');
            }
        }
    }

    /**
     * 检查浏览器兼容性
     * @returns {boolean} 是否兼容
     */
    function checkCompatibility() {
        // 检查必要的API支持
        const requiredAPIs = {
            FileReader: !!window.FileReader,
            Blob: !!window.Blob,
            URL: !!window.URL,
            localStorage: !!window.localStorage,
            JSON: !!window.JSON
        };

        const unsupportedAPIs = Object.keys(requiredAPIs).filter(api => !requiredAPIs[api]);

        if (unsupportedAPIs.length > 0) {
            const message = `浏览器不支持以下功能: ${unsupportedAPIs.join(', ')}。请升级浏览器。`;
            showFatalError(message);
            return false;
        }

        // 检查XLSX库是否加载
        if (!window.XLSX) {
            showFatalError('Excel解析库加载失败，请检查网络连接');
            return false;
        }

        return true;
    }

    /**
     * 按顺序初始化所有模块
     */
    function initializeModules() {
        const modules = [
            { name: 'Version', module: App.Version },
            { name: 'Utils', module: App.Utils },
            { name: 'FileHandler', module: App.FileHandler },
            { name: 'DataParser', module: App.DataParser },
            { name: 'TableRenderer', module: App.TableRenderer },
            { name: 'ExportUtils', module: App.ExportUtils },
            { name: 'ThemeManager', module: App.ThemeManager }
        ];

        modules.forEach(({ name, module }) => {
            if (module && typeof module.init === 'function') {
                try {
                    module.init();
                    console.log(`${name}模块初始化成功`);
                } catch (error) {
                    console.error(`${name}模块初始化失败:`, error);
                    throw new Error(`${name}模块初始化失败`);
                }
            } else {
                console.error(`${name}模块未找到或缺少init方法`);
                throw new Error(`${name}模块未找到`);
            }
        });

        // 显示版本更新通知
        setTimeout(() => {
            if (App.Version && typeof App.Version.showUpdateNotification === 'function') {
                App.Version.showUpdateNotification();
            }
        }, 1000);
    }

    /**
     * 设置全局事件监听器
     */
    function setupGlobalEventListeners() {
        // 页面卸载前的清理工作
        window.addEventListener('beforeunload', handleBeforeUnload);

        // 全局错误处理
        window.addEventListener('error', handleGlobalError);

        // 未处理的Promise拒绝
        window.addEventListener('unhandledrejection', handleUnhandledRejection);

        // 页面可见性变化
        document.addEventListener('visibilitychange', handleVisibilityChange);
    }

    /**
     * 设置全局函数（供HTML中的onclick等调用）
     */
    function setupGlobalFunctions() {
        // 工作表选择函数
        window.displaySheet = function() {
            if (App.TableRenderer && App.TableRenderer.isReady()) {
                const sheetSelect = document.getElementById('sheetSelect');
                if (sheetSelect && sheetSelect.value) {
                    App.TableRenderer.displaySheet(sheetSelect.value);
                }
            }
        };

        // 数据导出函数
        window.exportData = function(format) {
            if (App.ExportUtils && App.ExportUtils.isReady()) {
                App.ExportUtils.exportData(format);
            } else {
                console.error('导出模块未就绪');
                if (App.Utils) {
                    App.Utils.showStatus('导出功能暂不可用', 'error');
                }
            }
        };

        // 导出所有工作表函数
        window.exportAllSheets = function(format = 'json') {
            if (App.ExportUtils && App.ExportUtils.isReady()) {
                App.ExportUtils.exportAllSheets(format);
            }
        };
    }

    /**
     * 处理页面卸载前事件
     */
    function handleBeforeUnload() {
        // 清理资源
        if (App.DataParser) {
            App.DataParser.clearData();
        }
    }

    /**
     * 处理全局错误
     * @param {ErrorEvent} event - 错误事件
     */
    function handleGlobalError(event) {
        console.error('全局错误:', event.error);
        
        if (App.Utils) {
            App.Utils.showStatus('发生未知错误，请刷新页面重试', 'error');
        }
    }

    /**
     * 处理未捕获的Promise拒绝
     * @param {PromiseRejectionEvent} event - Promise拒绝事件
     */
    function handleUnhandledRejection(event) {
        console.error('未处理的Promise拒绝:', event.reason);
        
        if (App.Utils) {
            App.Utils.showStatus('操作失败，请重试', 'error');
        }
        
        // 阻止默认的控制台错误输出
        event.preventDefault();
    }

    /**
     * 处理页面可见性变化
     */
    function handleVisibilityChange() {
        if (document.hidden) {
            console.log('页面已隐藏');
        } else {
            console.log('页面已显示');
        }
    }

    /**
     * 显示致命错误
     * @param {string} message - 错误消息
     */
    function showFatalError(message) {
        console.error('致命错误:', message);
        
        // 尝试显示错误消息
        const uploadArea = document.getElementById('uploadArea');
        if (uploadArea) {
            uploadArea.innerHTML = `
                <div style="color: #dc3545; text-align: center; padding: 20px;">
                    <h3>⚠️ 应用程序错误</h3>
                    <p>${message}</p>
                    <button onclick="location.reload()" style="margin-top: 10px; padding: 8px 16px; background: #dc3545; color: white; border: none; border-radius: 4px; cursor: pointer;">
                        刷新页面
                    </button>
                </div>
            `;
            uploadArea.style.pointerEvents = 'auto';
        }
    }

    /**
     * 重置应用程序
     * 清除所有数据和状态
     */
    function reset() {
        try {
            // 重置各个模块
            if (App.FileHandler && App.FileHandler.reset) {
                App.FileHandler.reset();
            }
            
            if (App.DataParser && App.DataParser.clearData) {
                App.DataParser.clearData();
            }
            
            if (App.TableRenderer && App.TableRenderer.reset) {
                App.TableRenderer.reset();
            }

            if (App.ThemeManager && App.ThemeManager.resetAll) {
                App.ThemeManager.resetAll();
            }

            if (App.Utils) {
                App.Utils.showStatus('应用程序已重置', 'info');
            }
            
            console.log('应用程序已重置');
        } catch (error) {
            console.error('重置应用程序时出错:', error);
        }
    }

    /**
     * 获取应用程序状态
     * @returns {Object} 应用程序状态信息
     */
    function getStatus() {
        return {
            isInitialized: isInitialized,
            initializationAttempts: initializationAttempts,
            modules: {
                Utils: !!(App.Utils && App.Utils.isReady && App.Utils.isReady()),
                FileHandler: !!(App.FileHandler && App.FileHandler.isReady && App.FileHandler.isReady()),
                DataParser: !!(App.DataParser && App.DataParser.isReady && App.DataParser.isReady()),
                TableRenderer: !!(App.TableRenderer && App.TableRenderer.isReady && App.TableRenderer.isReady()),
                ExportUtils: !!(App.ExportUtils && App.ExportUtils.isReady && App.ExportUtils.isReady()),
                ThemeManager: !!(App.ThemeManager && App.ThemeManager.isReady && App.ThemeManager.isReady())
            }
        };
    }

    // 暴露公共接口
    return {
        init: init,
        reset: reset,
        getStatus: getStatus
    };

})();

// 页面加载完成后自动初始化应用程序
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM加载完成，开始初始化应用程序');
    App.Main.init();
});

// 模块加载完成日志
console.log('Main模块已加载 - 应用程序主控制器已就绪');
