# CORS问题解决方案

## 🚨 问题描述

当您直接双击打开`index.html`文件时，可能会遇到以下错误：

```
Access to fetch at 'file:///...' from origin 'null' has been blocked by CORS policy
```

这是因为现代浏览器出于安全考虑，不允许`file://`协议的页面访问其他本地文件。

## ✅ 解决方案

### 方案1：使用便捷启动脚本（推荐）

1. **双击运行** `start_server.bat` 文件
2. 脚本会自动：
   - 启动HTTP服务器
   - 打开浏览器访问 `http://localhost:8000`
3. 享受完整功能！

### 方案2：手动启动HTTP服务器

1. **打开命令提示符**（在项目根目录）
2. **运行命令**：
   ```bash
   python -m http.server 8000
   ```
3. **打开浏览器**访问：`http://localhost:8000`

### 方案3：使用开发工具

如果您使用VS Code等编辑器：

1. **安装Live Server扩展**
2. **右键点击** `index.html`
3. **选择** "Open with Live Server"

### 方案4：使用其他HTTP服务器

#### Node.js用户：
```bash
npx http-server -p 8000
```

#### PHP用户：
```bash
php -S localhost:8000
```

## 🎯 为什么需要HTTP服务器？

- **CORS限制**：浏览器安全策略限制`file://`协议访问其他文件
- **完整功能**：HTTP服务器环境下可以正常读取JSON映射文件
- **最佳体验**：避免功能受限，获得完整的工具体验

## 💡 推荐工作流程

1. **开发时**：使用`start_server.bat`或Live Server
2. **分享时**：告知用户使用HTTP服务器访问
3. **部署时**：部署到Web服务器上

## 🔧 故障排除

### 端口被占用
如果8000端口被占用，可以使用其他端口：
```bash
python -m http.server 8080
```

### Python未安装
- 下载安装Python：https://python.org
- 或使用其他HTTP服务器工具

### 防火墙问题
- 允许Python或HTTP服务器通过防火墙
- 或使用其他端口号

## 📋 功能对比

| 访问方式 | JSON映射 | 直接映射 | 文件保存 | 完整功能 |
|---------|---------|---------|---------|---------|
| file:// | ❌ 受限 | ✅ 正常 | ✅ 正常 | ⚠️ 部分 |
| HTTP服务器 | ✅ 正常 | ✅ 正常 | ✅ 正常 | ✅ 完整 |

**建议**：始终使用HTTP服务器访问以获得最佳体验！
