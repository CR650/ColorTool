/**
 * 增强版文件管理器
 * 结合多种浏览器API实现文件处理
 */

class EnhancedFileManager {
    constructor() {
        this.fileHandle = null;
        this.originalFile = null;
        this.capabilities = this.detectCapabilities();
    }

    /**
     * 检测浏览器能力
     */
    detectCapabilities() {
        return {
            fileSystemAccess: 'showOpenFilePicker' in window,
            fileAPI: 'File' in window,
            webkitDirectory: 'webkitdirectory' in document.createElement('input'),
            dragAndDrop: 'ondrop' in window
        };
    }

    /**
     * 智能文件选择（根据浏览器能力选择最佳方案）
     */
    async selectFile() {
        if (this.capabilities.fileSystemAccess) {
            return await this.selectWithFileSystemAPI();
        } else {
            return await this.selectWithFileInput();
        }
    }

    /**
     * 使用File System Access API选择文件
     */
    async selectWithFileSystemAPI() {
        try {
            const [fileHandle] = await window.showOpenFilePicker({
                types: [{
                    description: 'Excel files',
                    accept: {
                        'application/vnd.ms-excel': ['.xls'],
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
                    }
                }],
                multiple: false
            });

            this.fileHandle = fileHandle;
            this.originalFile = await fileHandle.getFile();
            
            return {
                method: 'fileSystemAccess',
                file: this.originalFile,
                canOverwrite: true
            };
        } catch (error) {
            throw new Error('文件选择失败: ' + error.message);
        }
    }

    /**
     * 使用传统文件输入选择文件
     */
    async selectWithFileInput() {
        return new Promise((resolve, reject) => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.xls,.xlsx';
            
            input.onchange = (event) => {
                const file = event.target.files[0];
                if (file) {
                    this.originalFile = file;
                    resolve({
                        method: 'fileInput',
                        file: file,
                        canOverwrite: false
                    });
                } else {
                    reject(new Error('未选择文件'));
                }
            };
            
            input.click();
        });
    }

    /**
     * 保存文件（智能选择保存方式）
     */
    async saveFile(data, options = {}) {
        if (this.capabilities.fileSystemAccess && this.fileHandle) {
            return await this.saveWithFileSystemAPI(data);
        } else {
            return await this.saveWithDownload(data, options);
        }
    }

    /**
     * 使用File System Access API直接覆盖文件
     */
    async saveWithFileSystemAPI(workbook) {
        try {
            // 请求写入权限
            const permission = await this.fileHandle.requestPermission({ mode: 'readwrite' });
            if (permission !== 'granted') {
                throw new Error('无法获取文件写入权限');
            }

            // 生成Excel数据 (使用xls格式以兼容Unity工具)
            const excelBuffer = XLSX.write(workbook, {
                bookType: 'xls',
                type: 'array'
            });

            // 创建可写流
            const writable = await this.fileHandle.createWritable();
            
            // 写入数据
            await writable.write(excelBuffer);
            await writable.close();

            return {
                success: true,
                method: 'fileSystemAccess',
                message: '文件已直接保存到原位置'
            };
        } catch (error) {
            throw new Error('文件保存失败: ' + error.message);
        }
    }

    /**
     * 使用下载方式保存文件
     */
    async saveWithDownload(workbook, options = {}) {
        try {
            // 生成Excel数据 (使用xls格式以兼容Unity工具)
            const excelBuffer = XLSX.write(workbook, {
                bookType: 'xls',
                type: 'array'
            });

            // 创建Blob (使用xls格式的MIME类型)
            const blob = new Blob([excelBuffer], {
                type: 'application/vnd.ms-excel'
            });

            // 创建下载链接
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = options.fileName || 'RSC_Theme_Updated.xls';
            
            // 触发下载
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            
            // 清理URL
            URL.revokeObjectURL(url);

            return {
                success: true,
                method: 'download',
                message: '文件已下载，请手动替换原文件'
            };
        } catch (error) {
            throw new Error('文件下载失败: ' + error.message);
        }
    }

    /**
     * 读取文件内容
     */
    async readFile() {
        if (!this.originalFile) {
            throw new Error('未选择文件');
        }

        const arrayBuffer = await this.originalFile.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1, 
            defval: '',
            raw: false 
        });

        return {
            workbook: workbook,
            data: jsonData,
            fileName: this.originalFile.name
        };
    }

    /**
     * 获取文件信息
     */
    getFileInfo() {
        if (!this.originalFile) {
            return null;
        }

        return {
            name: this.originalFile.name,
            size: this.originalFile.size,
            type: this.originalFile.type,
            lastModified: new Date(this.originalFile.lastModified),
            canOverwrite: !!this.fileHandle
        };
    }

    /**
     * 显示浏览器能力信息
     */
    getCapabilityInfo() {
        const info = [];
        
        if (this.capabilities.fileSystemAccess) {
            info.push('✅ 支持直接文件覆盖 (File System Access API)');
        } else {
            info.push('❌ 不支持直接文件覆盖，将使用下载方式');
        }
        
        if (this.capabilities.fileAPI) {
            info.push('✅ 支持文件读取 (File API)');
        }
        
        if (this.capabilities.dragAndDrop) {
            info.push('✅ 支持拖拽文件');
        }

        return info;
    }
}

// 导出实例
window.EnhancedFileManager = new EnhancedFileManager();
