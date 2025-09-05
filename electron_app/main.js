/**
 * Electron主进程
 * 桌面应用程序实现方案
 */

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs').promises;
const XLSX = require('xlsx');

class ElectronThemeManager {
    constructor() {
        this.mainWindow = null;
    }

    createWindow() {
        // 创建浏览器窗口
        this.mainWindow = new BrowserWindow({
            width: 1200,
            height: 800,
            webPreferences: {
                nodeIntegration: false,
                contextIsolation: true,
                preload: path.join(__dirname, 'preload.js')
            }
        });

        // 加载应用
        this.mainWindow.loadFile('index.html');

        // 开发模式下打开开发者工具
        if (process.env.NODE_ENV === 'development') {
            this.mainWindow.webContents.openDevTools();
        }
    }

    setupIPC() {
        // 选择RSC文件
        ipcMain.handle('select-rsc-file', async () => {
            const result = await dialog.showOpenDialog(this.mainWindow, {
                properties: ['openFile'],
                filters: [
                    { name: 'Excel Files', extensions: ['xls', 'xlsx'] }
                ]
            });

            if (!result.canceled && result.filePaths.length > 0) {
                return result.filePaths[0];
            }
            return null;
        });

        // 选择源数据文件
        ipcMain.handle('select-source-file', async () => {
            const result = await dialog.showOpenDialog(this.mainWindow, {
                properties: ['openFile'],
                filters: [
                    { name: 'Excel Files', extensions: ['xls', 'xlsx'] }
                ]
            });

            if (!result.canceled && result.filePaths.length > 0) {
                return result.filePaths[0];
            }
            return null;
        });

        // 读取Excel文件
        ipcMain.handle('read-excel-file', async (event, filePath, sheetName = null) => {
            try {
                const fileBuffer = await fs.readFile(filePath);
                const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
                
                let worksheet;
                if (sheetName && workbook.SheetNames.includes(sheetName)) {
                    worksheet = workbook.Sheets[sheetName];
                } else {
                    worksheet = workbook.Sheets[workbook.SheetNames[0]];
                }

                const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                    header: 1,
                    defval: '',
                    raw: false
                });

                return {
                    success: true,
                    data: jsonData,
                    sheetNames: workbook.SheetNames,
                    fileName: path.basename(filePath)
                };
            } catch (error) {
                return {
                    success: false,
                    error: error.message
                };
            }
        });

        // 直接覆盖保存Excel文件
        ipcMain.handle('save-excel-file', async (event, filePath, data) => {
            try {
                // 创建新的工作簿
                const newWorkbook = XLSX.utils.book_new();
                const newWorksheet = XLSX.utils.aoa_to_sheet(data);
                XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');

                // 生成文件缓冲区 (使用xls格式以兼容Unity工具)
                const fileBuffer = XLSX.write(newWorkbook, {
                    type: 'buffer',
                    bookType: 'xls'
                });

                // 创建备份 (保持xls格式)
                const backupPath = filePath.replace(/\.xlsx?$/, '_backup_' + Date.now() + '.xls');
                try {
                    const originalBuffer = await fs.readFile(filePath);
                    await fs.writeFile(backupPath, originalBuffer);
                    console.log('备份文件已创建:', backupPath);
                } catch (backupError) {
                    console.warn('创建备份失败:', backupError.message);
                }

                // 覆盖原文件
                await fs.writeFile(filePath, fileBuffer);

                return {
                    success: true,
                    message: '文件已成功保存',
                    backupPath: backupPath
                };
            } catch (error) {
                return {
                    success: false,
                    error: error.message
                };
            }
        });

        // 显示保存对话框
        ipcMain.handle('show-save-dialog', async (event, defaultName) => {
            const result = await dialog.showSaveDialog(this.mainWindow, {
                defaultPath: defaultName,
                filters: [
                    { name: 'Excel Files', extensions: ['xlsx'] }
                ]
            });

            if (!result.canceled) {
                return result.filePath;
            }
            return null;
        });

        // 显示消息框
        ipcMain.handle('show-message', async (event, options) => {
            return await dialog.showMessageBox(this.mainWindow, options);
        });

        // 获取文件信息
        ipcMain.handle('get-file-info', async (event, filePath) => {
            try {
                const stats = await fs.stat(filePath);
                return {
                    success: true,
                    info: {
                        size: stats.size,
                        modified: stats.mtime,
                        name: path.basename(filePath),
                        directory: path.dirname(filePath)
                    }
                };
            } catch (error) {
                return {
                    success: false,
                    error: error.message
                };
            }
        });
    }

    init() {
        // 应用准备就绪时创建窗口
        app.whenReady().then(() => {
            this.createWindow();
            this.setupIPC();

            app.on('activate', () => {
                if (BrowserWindow.getAllWindows().length === 0) {
                    this.createWindow();
                }
            });
        });

        // 所有窗口关闭时退出应用
        app.on('window-all-closed', () => {
            if (process.platform !== 'darwin') {
                app.quit();
            }
        });
    }
}

// 创建并初始化应用
const themeManager = new ElectronThemeManager();
themeManager.init();
