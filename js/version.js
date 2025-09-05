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
        current: '1.1.0',
        releaseDate: '2025-09-05',
        buildNumber: '20250905001'
    };

    // 版本历史记录
    const VERSION_HISTORY = [
        {
            version: '1.1.0',
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
