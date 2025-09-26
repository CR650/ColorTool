/**
 * Unity项目文件夹管理器
 * 文件说明：实现文件夹选择和自动文件定位功能，简化用户操作流程
 * 创建时间：2025-09-05
 * 依赖：File System Access API, XLSX库
 */

// 确保全局命名空间存在
window.App = window.App || {};

/**
 * Unity项目文件夹管理器
 * 负责文件夹选择、文件自动定位和权限管理
 */
window.App.UnityProjectFolderManager = (function() {
    'use strict';

    /**
     * 文件夹管理器类
     */
    class UnityProjectFolderManager {
        constructor() {
            this.directoryHandle = null;
            this.rscThemeHandle = null;
            this.ugcThemeHandle = null;
            this.rscLanguageHandle = null;
            this.levelsHandle = null;
            this.allObstacleHandle = null;
            this.selectedFolderPath = null; // 存储选择的文件夹路径信息
            this.isSupported = 'showDirectoryPicker' in window;
            this.cache = new Map();
            this.maxCacheAge = 5 * 60 * 1000; // 5分钟缓存
        }

        /**
         * 检查浏览器支持
         */
        checkSupport() {
            if (!this.isSupported) {
                throw new Error('当前浏览器不支持文件夹选择功能，请使用Chrome 86+或Edge 86+');
            }
            return true;
        }

        /**
         * 选择Unity项目文件夹并自动定位主题文件
         */
        async selectUnityProjectFolder() {
            this.checkSupport();

            try {
                console.log('开始选择Unity项目文件夹...');
                
                // 选择文件夹
                this.directoryHandle = await window.showDirectoryPicker({
                    id: 'unity-xlsx-folder',
                    mode: 'readwrite',
                    startIn: 'documents'
                });

                console.log('文件夹选择成功:', this.directoryHandle.name);

                // 存储文件夹路径信息（用于后续路径显示）
                this.selectedFolderPath = this.directoryHandle.name;

                // 验证文件夹路径
                const isValidFolder = await this.validateUnityFolder();
                if (!isValidFolder) {
                    throw new Error('请选择Unity项目的Tools/xlsx文件夹，或包含RSC_Theme.xls/UGCTheme.xls的文件夹');
                }

                // 自动定位主题文件
                const locatedFiles = await this.locateThemeFiles();

                const result = {
                    success: true,
                    directoryPath: this.directoryHandle.name,
                    selectedFolderPath: this.selectedFolderPath, // 添加路径信息
                    rscThemeFound: !!locatedFiles.rscTheme,
                    ugcThemeFound: !!locatedFiles.ugcTheme,
                    rscLanguageFound: !!locatedFiles.rscLanguage,
                    levelsFound: !!locatedFiles.levels,
                    allObstacleFound: !!locatedFiles.allObstacle,
                    files: locatedFiles
                };

                console.log('文件夹选择结果:', result);
                return result;

            } catch (error) {
                console.error('文件夹选择失败:', error);
                throw error;
            }
        }

        /**
         * 验证是否为Unity项目的xlsx文件夹
         */
        async validateUnityFolder() {
            try {
                console.log('验证文件夹:', this.directoryHandle.name);

                // 检查是否包含预期的文件
                const expectedFiles = ['RSC_Theme.xls', 'UGCTheme.xls', 'RSC_Language.xls', 'AllObstacle.xls'];
                let foundCount = 0;
                const foundFiles = [];

                for await (const [name, handle] of this.directoryHandle.entries()) {
                    if (expectedFiles.includes(name)) {
                        foundCount++;
                        foundFiles.push(name);
                    }
                }

                console.log(`找到 ${foundCount} 个主题文件:`, foundFiles);
                return foundCount >= 1; // 至少找到一个主题文件

            } catch (error) {
                console.error('文件夹验证失败:', error);
                return false;
            }
        }

        /**
         * 自动定位主题文件
         */
        async locateThemeFiles() {
            const result = {
                rscTheme: null,
                ugcTheme: null,
                rscLanguage: null,
                levels: null,
                allObstacle: null
            };

            try {
                console.log('开始定位主题文件...');

                for await (const [name, handle] of this.directoryHandle.entries()) {
                    if (handle.kind === 'file') {
                        if (name === 'RSC_Theme.xls') {
                            console.log('找到RSC_Theme.xls，请求权限...');
                            const permission = await handle.requestPermission({ mode: 'readwrite' });
                            if (permission === 'granted') {
                                this.rscThemeHandle = handle;
                                result.rscTheme = {
                                    handle: handle,
                                    name: name,
                                    hasPermission: true
                                };
                                console.log('RSC_Theme.xls权限获取成功');
                            } else {
                                console.warn('RSC_Theme.xls权限获取失败');
                                result.rscTheme = {
                                    handle: handle,
                                    name: name,
                                    hasPermission: false
                                };
                            }
                        } else if (name === 'UGCTheme.xls') {
                            console.log('找到UGCTheme.xls，请求权限...');
                            const permission = await handle.requestPermission({ mode: 'readwrite' });
                            if (permission === 'granted') {
                                this.ugcThemeHandle = handle;
                                result.ugcTheme = {
                                    handle: handle,
                                    name: name,
                                    hasPermission: true
                                };
                                console.log('UGCTheme.xls权限获取成功');
                            } else {
                                console.warn('UGCTheme.xls权限获取失败');
                                result.ugcTheme = {
                                    handle: handle,
                                    name: name,
                                    hasPermission: false
                                };
                            }
                        } else if (name === 'RSC_Language.xls') {
                            console.log('找到RSC_Language.xls，请求权限...');
                            const permission = await handle.requestPermission({ mode: 'readwrite' });
                            if (permission === 'granted') {
                                this.rscLanguageHandle = handle;
                                result.rscLanguage = {
                                    handle: handle,
                                    name: name,
                                    hasPermission: true
                                };
                                console.log('RSC_Language.xls权限获取成功');
                            } else {
                                console.warn('RSC_Language.xls权限获取失败');
                                result.rscLanguage = {
                                    handle: handle,
                                    name: name,
                                    hasPermission: false
                                };
                            }
                        } else if (name === 'Levels.xls') {
                            console.log('找到Levels.xls，请求权限...');
                            const permission = await handle.requestPermission({ mode: 'read' });
                            if (permission === 'granted') {
                                this.levelsHandle = handle;
                                result.levels = {
                                    handle: handle,
                                    name: name,
                                    hasPermission: true
                                };
                                console.log('Levels.xls权限获取成功');
                            } else {
                                console.warn('Levels.xls权限获取失败');
                                result.levels = {
                                    handle: handle,
                                    name: name,
                                    hasPermission: false
                                };
                            }
                        } else if (name === 'AllObstacle.xls') {
                            console.log('找到AllObstacle.xls，请求权限...');
                            const permission = await handle.requestPermission({ mode: 'readwrite' });
                            if (permission === 'granted') {
                                this.allObstacleHandle = handle;
                                result.allObstacle = {
                                    handle: handle,
                                    name: name,
                                    hasPermission: true
                                };
                                console.log('AllObstacle.xls权限获取成功');
                            } else {
                                console.warn('AllObstacle.xls权限获取失败');
                                result.allObstacle = {
                                    handle: handle,
                                    name: name,
                                    hasPermission: false
                                };
                            }
                        }
                    }
                }

                console.log('文件定位完成:', result);
                return result;

            } catch (error) {
                console.error('文件定位失败:', error);
                throw error;
            }
        }

        /**
         * 读取主题文件数据
         */
        async loadThemeFileData(fileType) {
            let handle;
            if (fileType === 'rsc') {
                handle = this.rscThemeHandle;
            } else if (fileType === 'ugc') {
                handle = this.ugcThemeHandle;
            } else if (fileType === 'allObstacle') {
                handle = this.allObstacleHandle;
            } else {
                throw new Error(`不支持的文件类型: ${fileType}`);
            }

            if (!handle) {
                throw new Error(`${fileType.toUpperCase()}文件未找到`);
            }

            try {
                console.log(`开始读取${fileType.toUpperCase()}文件数据...`);
                
                const file = await handle.getFile();
                const arrayBuffer = await file.arrayBuffer();
                
                // 使用XLSX解析
                const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                
                const result = {
                    workbook: workbook,
                    fileHandle: handle,
                    fileName: file.name,
                    fileSize: file.size,
                    lastModified: new Date(file.lastModified)
                };

                console.log(`${fileType.toUpperCase()}文件读取成功:`, {
                    fileName: result.fileName,
                    fileSize: result.fileSize,
                    sheetCount: workbook.SheetNames.length
                });

                return result;

            } catch (error) {
                console.error(`读取${fileType}文件失败:`, error);
                throw error;
            }
        }

        /**
         * 获取文件夹信息
         */
        getFolderInfo() {
            if (!this.directoryHandle) {
                return null;
            }

            return {
                name: this.directoryHandle.name,
                kind: this.directoryHandle.kind,
                hasRSCTheme: !!this.rscThemeHandle,
                hasUGCTheme: !!this.ugcThemeHandle,
                hasAllObstacle: !!this.allObstacleHandle
            };
        }

        /**
         * 清理资源
         */
        cleanup() {
            this.directoryHandle = null;
            this.rscThemeHandle = null;
            this.ugcThemeHandle = null;
            this.rscLanguageHandle = null;
            this.levelsHandle = null;
            this.allObstacleHandle = null;
            this.cache.clear();
        }
    }

    // 暴露公共接口
    return {
        UnityProjectFolderManager: UnityProjectFolderManager,
        
        // 创建实例的便捷方法
        create: function() {
            return new UnityProjectFolderManager();
        },
        
        // 检查浏览器支持
        isSupported: function() {
            return 'showDirectoryPicker' in window;
        }
    };

})();
