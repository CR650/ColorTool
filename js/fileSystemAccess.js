/**
 * File System Access API 实现方案
 * 直接修改本地RSC_Theme.xls文件
 */

class FileSystemManager {
    constructor() {
        this.fileHandle = null;
        this.isSupported = 'showOpenFilePicker' in window;
    }

    /**
     * 检查浏览器支持
     */
    checkSupport() {
        if (!this.isSupported) {
            throw new Error('当前浏览器不支持File System Access API。请使用Chrome 86+、Edge 86+或其他支持的浏览器。');
        }
        return true;
    }

    /**
     * 选择并获取RSC_Theme文件的句柄
     */
    async selectRSCThemeFile() {
        try {
            this.checkSupport();

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
            console.log('文件句柄获取成功:', fileHandle.name);
            
            // 验证文件权限
            const permission = await this.verifyPermission(fileHandle, true);
            if (!permission) {
                throw new Error('无法获取文件写入权限');
            }

            return fileHandle;
        } catch (error) {
            console.error('选择文件失败:', error);
            throw error;
        }
    }

    /**
     * 验证文件权限
     */
    async verifyPermission(fileHandle, readWrite = false) {
        const options = {};
        if (readWrite) {
            options.mode = 'readwrite';
        }

        // 检查是否已有权限
        if ((await fileHandle.queryPermission(options)) === 'granted') {
            return true;
        }

        // 请求权限
        if ((await fileHandle.requestPermission(options)) === 'granted') {
            return true;
        }

        return false;
    }

    /**
     * 读取文件内容
     */
    async readFile() {
        if (!this.fileHandle) {
            throw new Error('未选择文件');
        }

        const file = await this.fileHandle.getFile();
        const arrayBuffer = await file.arrayBuffer();
        
        // 使用XLSX解析
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
            fileName: file.name
        };
    }

    /**
     * 直接覆盖原文件
     */
    async writeFile(updatedWorkbook) {
        if (!this.fileHandle) {
            throw new Error('未选择文件');
        }

        try {
            // 生成Excel文件的二进制数据
            const excelBuffer = XLSX.write(updatedWorkbook, {
                bookType: 'xlsx',
                type: 'array'
            });

            // 创建可写流
            const writable = await this.fileHandle.createWritable();
            
            // 写入数据
            await writable.write(excelBuffer);
            
            // 关闭流
            await writable.close();

            console.log('文件已成功覆盖保存');
            return true;
        } catch (error) {
            console.error('文件写入失败:', error);
            throw error;
        }
    }

    /**
     * 获取文件信息
     */
    getFileInfo() {
        if (!this.fileHandle) {
            return null;
        }

        return {
            name: this.fileHandle.name,
            kind: this.fileHandle.kind
        };
    }
}

// 导出实例
window.FileSystemManager = new FileSystemManager();
