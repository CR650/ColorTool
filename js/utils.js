/**
 * 通用工具函数模块
 * 文件说明：包含项目中使用的通用工具函数，如状态提示、文件大小格式化等
 * 创建时间：2025-01-09
 * 依赖：无
 */

// 创建全局命名空间
window.App = window.App || {};

/**
 * 工具函数模块
 * 提供通用的工具函数和辅助方法
 */
window.App.Utils = (function() {
    'use strict';

    // 状态消息DOM元素引用
    let statusMessageElement = null;

    /**
     * 初始化工具模块
     * 获取必要的DOM元素引用
     */
    function init() {
        statusMessageElement = document.getElementById('statusMessage');
    }

    /**
     * 显示状态消息
     * @param {string} message - 要显示的消息内容
     * @param {string} type - 消息类型：'info', 'success', 'error'
     * @param {number} duration - 显示持续时间（毫秒），默认3000ms
     */
    function showStatus(message, type = 'info', duration = 3000) {
        if (!statusMessageElement) {
            console.warn('状态消息元素未找到');
            return;
        }

        // 设置消息内容和样式
        statusMessageElement.textContent = message;
        statusMessageElement.className = `status-message ${type} show`;
        
        // 自动隐藏消息
        setTimeout(() => {
            statusMessageElement.classList.remove('show');
        }, duration);
    }

    /**
     * 格式化文件大小
     * 将字节数转换为人类可读的格式
     * @param {number} bytes - 文件大小（字节）
     * @returns {string} 格式化后的文件大小字符串
     */
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    /**
     * 验证文件类型
     * 检查文件是否为支持的表格格式
     * @param {File} file - 要验证的文件对象
     * @returns {boolean} 是否为支持的文件类型
     */
    function validateFileType(file) {
        // 支持的MIME类型
        const allowedTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
            'application/vnd.ms-excel', // .xls
            'text/csv', // .csv
            'application/csv'
        ];
        
        // 支持的文件扩展名
        const allowedExtensions = ['.xlsx', '.xls', '.csv'];
        const fileName = file.name.toLowerCase();
        const hasValidExtension = allowedExtensions.some(ext => fileName.endsWith(ext));
        
        // 通过MIME类型或文件扩展名验证
        return allowedTypes.includes(file.type) || hasValidExtension;
    }

    /**
     * 检查文件大小
     * 验证文件是否超过最大允许大小
     * @param {File} file - 要检查的文件对象
     * @param {number} maxSizeMB - 最大允许大小（MB），默认10MB
     * @returns {boolean} 文件大小是否在允许范围内
     */
    function validateFileSize(file, maxSizeMB = 10) {
        const maxSizeBytes = maxSizeMB * 1024 * 1024;
        return file.size <= maxSizeBytes;
    }

    /**
     * 防抖函数
     * 限制函数的执行频率，在指定时间内只执行最后一次调用
     * @param {Function} func - 要防抖的函数
     * @param {number} wait - 等待时间（毫秒）
     * @returns {Function} 防抖后的函数
     */
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

    /**
     * 节流函数
     * 限制函数的执行频率，在指定时间间隔内最多执行一次
     * @param {Function} func - 要节流的函数
     * @param {number} limit - 时间间隔（毫秒）
     * @returns {Function} 节流后的函数
     */
    function throttle(func, limit) {
        let inThrottle;
        return function(...args) {
            if (!inThrottle) {
                func.apply(this, args);
                inThrottle = true;
                setTimeout(() => inThrottle = false, limit);
            }
        };
    }

    /**
     * 深拷贝对象
     * 创建对象的深层副本
     * @param {any} obj - 要拷贝的对象
     * @returns {any} 深拷贝后的对象
     */
    function deepClone(obj) {
        if (obj === null || typeof obj !== 'object') {
            return obj;
        }
        
        if (obj instanceof Date) {
            return new Date(obj.getTime());
        }
        
        if (obj instanceof Array) {
            return obj.map(item => deepClone(item));
        }
        
        if (typeof obj === 'object') {
            const clonedObj = {};
            for (let key in obj) {
                if (obj.hasOwnProperty(key)) {
                    clonedObj[key] = deepClone(obj[key]);
                }
            }
            return clonedObj;
        }
    }

    /**
     * 生成唯一ID
     * 生成基于时间戳和随机数的唯一标识符
     * @returns {string} 唯一ID字符串
     */
    function generateUniqueId() {
        return Date.now().toString(36) + Math.random().toString(36).substr(2);
    }

    /**
     * 检查浏览器兼容性
     * 检查当前浏览器是否支持必要的API
     * @returns {object} 兼容性检查结果
     */
    function checkBrowserCompatibility() {
        const result = {
            fileReader: !!window.FileReader,
            xlsx: !!window.XLSX,
            dragAndDrop: 'draggable' in document.createElement('div'),
            localStorage: !!window.localStorage
        };

        result.isCompatible = Object.values(result).every(Boolean);

        return result;
    }

    // 文件路径记忆常量
    const PATH_MEMORY_KEYS = {
        SOURCE_DATA: 'colorTool_lastPath_source',
        RSC_THEME: 'colorTool_lastPath_rsc',
        UGC_THEME: 'colorTool_lastPath_ugc'
    };

    /**
     * 保存最后选择的文件信息
     * @param {string} fileType - 文件类型 ('SOURCE_DATA', 'RSC_THEME', 'UGC_THEME')
     * @param {string} fileName - 文件名
     */
    function saveLastPath(fileType, fileName) {
        try {
            if (!fileName || !PATH_MEMORY_KEYS[fileType]) {
                return;
            }

            // 保存文件名和时间戳，用于用户体验提升
            const fileInfo = {
                fileName: fileName,
                timestamp: Date.now()
            };

            localStorage.setItem(PATH_MEMORY_KEYS[fileType], JSON.stringify(fileInfo));
            console.log(`已保存${fileType}文件记忆:`, fileName);
        } catch (error) {
            console.warn('保存文件记忆失败:', error);
        }
    }

    /**
     * 获取最后选择的文件信息
     * @param {string} fileType - 文件类型 ('SOURCE_DATA', 'RSC_THEME', 'UGC_THEME')
     * @returns {Object|null} 文件信息对象或null
     */
    function getLastPath(fileType) {
        try {
            if (PATH_MEMORY_KEYS[fileType]) {
                const savedData = localStorage.getItem(PATH_MEMORY_KEYS[fileType]);
                if (savedData) {
                    const fileInfo = JSON.parse(savedData);
                    console.log(`获取${fileType}文件记忆:`, fileInfo.fileName);
                    return fileInfo;
                }
            }
        } catch (error) {
            console.warn('获取文件记忆失败:', error);
        }
        return null;
    }

    /**
     * 获取推荐的startIn位置
     * @param {string} fileType - 文件类型
     * @returns {string} 推荐的startIn位置
     */
    function getRecommendedStartIn(fileType) {
        // 根据文件类型返回合适的预定义位置
        switch (fileType) {
            case 'SOURCE_DATA':
                return 'documents'; // 源数据文件通常在文档目录
            case 'RSC_THEME':
            case 'UGC_THEME':
                return 'documents'; // 主题文件也通常在文档目录
            default:
                return 'documents';
        }
    }

    // 暴露公共接口
    return {
        init: init,
        showStatus: showStatus,
        formatFileSize: formatFileSize,
        validateFileType: validateFileType,
        validateFileSize: validateFileSize,
        debounce: debounce,
        throttle: throttle,
        deepClone: deepClone,
        generateUniqueId: generateUniqueId,
        checkBrowserCompatibility: checkBrowserCompatibility,
        saveLastPath: saveLastPath,
        getLastPath: getLastPath,
        getRecommendedStartIn: getRecommendedStartIn
    };

})();

// 模块加载完成日志
console.log('Utils模块已加载 - 通用工具函数已就绪');
