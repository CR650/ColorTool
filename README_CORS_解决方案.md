# CORS问题解决方案

## 问题描述
当直接在浏览器中打开HTML文件时（file://协议），会遇到CORS错误，无法读取JSON映射文件。

## 解决方案

### 方案1：使用HTTP服务器（推荐）

#### Windows用户：
1. 双击运行 `start_server.bat` 文件
2. 或在命令行中运行：
   ```cmd
   python -m http.server 8000
   ```
3. 在浏览器中访问：`http://localhost:8000`

#### 其他系统：
```bash
# Python 3
python -m http.server 8000

# Python 2
python -m SimpleHTTPServer 8000

# Node.js (需要先安装 http-server)
npx http-server -p 8000
```

### 方案2：使用开发工具
- **VS Code**: 安装 "Live Server" 扩展
- **WebStorm**: 右键HTML文件选择 "Open in Browser"
- **Chrome**: 启动时添加参数 `--disable-web-security --user-data-dir="C:/temp"`

### 方案3：部署到Web服务器
将项目文件上传到任何Web服务器（Apache、Nginx、IIS等）

## 验证成功
当看到控制台输出以下信息时，表示JSON文件加载成功：
```
✅ JSON映射数据验证通过，包含 32 个映射项
✨ 检测到扩展映射数据！支持32个颜色通道（超过默认的17个）
```

## 注意事项
- 确保 `XLS/对比.json` 文件存在且格式正确
- 如果JSON文件读取失败，工具会自动回退到内置的17个通道映射
- 建议始终通过HTTP服务器访问项目以获得最佳体验
