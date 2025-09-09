/**
 * 图案选择器模块
 * 文件说明：实现Floor_Pattern.png图片中单个图案的选择功能
 * 创建时间：2025-09-08
 */

// 确保全局命名空间存在
window.App = window.App || {};

/**
 * 图案选择器模块
 */
window.App.PatternSelector = (function() {
    'use strict';

    // 模块状态
    let isInitialized = false;
    let patternImage = null;
    let selectedPattern = null;
    let patternGrid = null;

    // DOM元素引用
    let patternContainer = null;
    let patternCanvas = null;
    let patternPreview = null;
    let patternInfo = null;

    // 图案网格配置（根据实际图片调整）
    const PATTERN_CONFIG = {
        imageWidth: 1024,
        imageHeight: 1024,
        gridCols: 8,
        gridRows: 8,
        patternWidth: 128,
        patternHeight: 128
    };

    /**
     * 初始化图案选择器
     */
    function init() {
        if (isInitialized) {
            console.warn('PatternSelector模块已经初始化');
            return;
        }

        console.log('初始化图案选择器模块...');

        // 获取DOM元素
        patternContainer = document.getElementById('patternSelectorContainer');
        patternCanvas = document.getElementById('patternCanvas');
        patternPreview = document.getElementById('patternPreview');
        patternInfo = document.getElementById('patternInfo');

        if (!patternContainer) {
            console.error('未找到图案选择器容器元素');
            return;
        }

        // 初始化图案网格
        initializePatternGrid();

        // 加载图案图片
        loadPatternImage();

        // 绑定事件
        bindEvents();

        isInitialized = true;
        console.log('✅ 图案选择器模块初始化完成');
    }

    /**
     * 初始化图案网格配置
     */
    function initializePatternGrid() {
        patternGrid = [];

        let index = 0;
        for (let row = 0; row < PATTERN_CONFIG.gridRows; row++) {
            for (let col = 0; col < PATTERN_CONFIG.gridCols; col++) {
                const pattern = {
                    id: `pattern_${row}_${col}`,
                    index: index,
                    row: row,
                    col: col,
                    x: col * PATTERN_CONFIG.patternWidth,
                    y: row * PATTERN_CONFIG.patternHeight,
                    width: PATTERN_CONFIG.patternWidth,
                    height: PATTERN_CONFIG.patternHeight,
                    name: `图案 ${index}`,
                    selectable: true // 默认可选择，将在图片加载后检测
                };
                patternGrid.push(pattern);
                index++;
            }
        }

        console.log(`初始化图案网格: ${PATTERN_CONFIG.gridRows}x${PATTERN_CONFIG.gridCols} = ${patternGrid.length}个图案`);
    }

    /**
     * 加载图案图片
     */
    function loadPatternImage() {
        console.log('开始加载图案图片...');
        
        patternImage = new Image();
        patternImage.onload = function() {
            console.log('图案图片加载成功:', {
                width: this.naturalWidth,
                height: this.naturalHeight
            });

            // 更新配置（如果实际尺寸不同）
            if (this.naturalWidth !== PATTERN_CONFIG.imageWidth || 
                this.naturalHeight !== PATTERN_CONFIG.imageHeight) {
                
                console.log('更新图案配置以匹配实际图片尺寸');
                PATTERN_CONFIG.imageWidth = this.naturalWidth;
                PATTERN_CONFIG.imageHeight = this.naturalHeight;
                PATTERN_CONFIG.patternWidth = Math.floor(this.naturalWidth / PATTERN_CONFIG.gridCols);
                PATTERN_CONFIG.patternHeight = Math.floor(this.naturalHeight / PATTERN_CONFIG.gridRows);
                
                // 重新初始化网格
                initializePatternGrid();
            }

            // 渲染图案选择器（延迟检测黑块以避免CORS问题）
            renderPatternSelector();
        };

        patternImage.onerror = function() {
            console.error('图案图片加载失败');
            showError('无法加载图案图片，请检查文件路径');
        };

        // 设置图片路径
        patternImage.src = './Texture/Floor_Pattern.png';
    }

    /**
     * 检测是否为纯黑图案（延迟检测以避免CORS问题）
     */
    function isBlackPattern(pattern) {
        try {
            // 创建临时canvas来检测图案颜色
            const tempCanvas = document.createElement('canvas');
            const tempCtx = tempCanvas.getContext('2d');
            const sampleSize = 32; // 使用较小的采样尺寸提高性能
            tempCanvas.width = sampleSize;
            tempCanvas.height = sampleSize;

            // 绘制图案到临时canvas
            tempCtx.drawImage(
                patternImage,
                pattern.x, pattern.y, pattern.width, pattern.height,
                0, 0, sampleSize, sampleSize
            );

            // 获取图像数据
            const imageData = tempCtx.getImageData(0, 0, sampleSize, sampleSize);
            const data = imageData.data;

            // 检查是否所有像素都是纯黑色
            let blackPixelCount = 0;
            const totalPixels = data.length / 4;
            const threshold = 10; // 更严格的黑色阈值，只有接近纯黑的才算

            for (let i = 0; i < data.length; i += 4) {
                const r = data[i];
                const g = data[i + 1];
                const b = data[i + 2];
                const alpha = data[i + 3];

                // 只有RGB值都非常低且不透明，才认为是黑色
                if (alpha > 200 && r <= threshold && g <= threshold && b <= threshold) {
                    blackPixelCount++;
                }
            }

            // 只有100%纯黑的图案才算黑块
            const blackRatio = blackPixelCount / totalPixels;
            const isBlack = blackRatio >= 0.99; // 99%以上才算纯黑（考虑到压缩等因素的微小误差）

            console.log(`图案 ${pattern.index} 黑色像素比例: ${(blackRatio * 100).toFixed(1)}%, 判定为黑块: ${isBlack}`);
            return isBlack;

        } catch (error) {
            // 如果遇到CORS错误，使用启发式方法
            console.warn(`无法检测图案 ${pattern.index} 的颜色（CORS限制）:`, error.message);
            return isBlackPatternHeuristic(pattern);
        }
    }

    /**
     * 启发式黑块检测（基于已知模式）
     */
    function isBlackPatternHeuristic(pattern) {
        // 根据观察到的图片，这些位置是纯黑块
        const knownBlackPositions = [
            // 第一行的黑块（除了index 0）
            6, 7,
            // 第二行的黑块
            8, 15,
            // 第三行的黑块
            23,
            // 第四行的黑块
            25, 26, 31,
            // 第五行的黑块
            32, 33, 34, 35, 36, 37, 38, 39,
            // 第六行的黑块
            40, 41, 42, 43, 44, 45, 46, 47,
            // 第七行的黑块
            48, 49, 50, 51, 52, 53, 54, 55,
            // 第八行的黑块
            56, 57, 58, 59, 60, 61, 62, 63
        ];

        return knownBlackPositions.includes(pattern.index);
    }

    /**
     * 渲染图案选择器
     */
    function renderPatternSelector() {
        if (!patternCanvas || !patternImage) {
            console.error('Canvas元素或图片未准备好');
            return;
        }

        console.log('开始渲染图案选择器...');

        // 设置Canvas尺寸
        const containerWidth = patternContainer.clientWidth || 800;
        const aspectRatio = PATTERN_CONFIG.imageHeight / PATTERN_CONFIG.imageWidth;
        const canvasWidth = Math.min(containerWidth - 40, 800); // 留出边距
        const canvasHeight = canvasWidth * aspectRatio;

        patternCanvas.width = canvasWidth;
        patternCanvas.height = canvasHeight;
        patternCanvas.style.width = canvasWidth + 'px';
        patternCanvas.style.height = canvasHeight + 'px';

        // 获取绘图上下文
        const ctx = patternCanvas.getContext('2d');

        // 绘制图案图片
        ctx.drawImage(patternImage, 0, 0, canvasWidth, canvasHeight);

        // 绘制网格线（可选）
        drawGrid(ctx, canvasWidth, canvasHeight);

        // 绘制不可选择图案的遮罩
        drawUnselectableMask(ctx, canvasWidth, canvasHeight);

        console.log('✅ 图案选择器渲染完成');
    }

    /**
     * 绘制网格线
     */
    function drawGrid(ctx, canvasWidth, canvasHeight) {
        ctx.strokeStyle = 'rgba(255, 255, 255, 0.3)';
        ctx.lineWidth = 1;

        const cellWidth = canvasWidth / PATTERN_CONFIG.gridCols;
        const cellHeight = canvasHeight / PATTERN_CONFIG.gridRows;

        // 绘制垂直线
        for (let i = 1; i < PATTERN_CONFIG.gridCols; i++) {
            const x = i * cellWidth;
            ctx.beginPath();
            ctx.moveTo(x, 0);
            ctx.lineTo(x, canvasHeight);
            ctx.stroke();
        }

        // 绘制水平线
        for (let i = 1; i < PATTERN_CONFIG.gridRows; i++) {
            const y = i * cellHeight;
            ctx.beginPath();
            ctx.moveTo(0, y);
            ctx.lineTo(canvasWidth, y);
            ctx.stroke();
        }
    }

    /**
     * 绘制不可选择图案的遮罩
     */
    function drawUnselectableMask(ctx, canvasWidth, canvasHeight) {
        const cellWidth = canvasWidth / PATTERN_CONFIG.gridCols;
        const cellHeight = canvasHeight / PATTERN_CONFIG.gridRows;

        patternGrid.forEach(pattern => {
            // 只对已检测过且不可选择的图案绘制遮罩
            if (pattern.blackChecked && !pattern.selectable) {
                const x = pattern.col * cellWidth;
                const y = pattern.row * cellHeight;

                // 绘制半透明红色遮罩
                ctx.fillStyle = 'rgba(220, 53, 69, 0.6)';
                ctx.fillRect(x, y, cellWidth, cellHeight);

                // 绘制禁用图标（X）
                ctx.strokeStyle = 'rgba(255, 255, 255, 0.9)';
                ctx.lineWidth = 3;
                const centerX = x + cellWidth / 2;
                const centerY = y + cellHeight / 2;
                const size = Math.min(cellWidth, cellHeight) * 0.3;

                // 绘制X
                ctx.beginPath();
                ctx.moveTo(centerX - size, centerY - size);
                ctx.lineTo(centerX + size, centerY + size);
                ctx.moveTo(centerX + size, centerY - size);
                ctx.lineTo(centerX - size, centerY + size);
                ctx.stroke();
            }
        });
    }

    /**
     * 绑定事件
     */
    function bindEvents() {
        if (!patternCanvas) return;

        // Canvas点击事件
        patternCanvas.addEventListener('click', handleCanvasClick);
        
        // Canvas鼠标移动事件（显示悬停效果）
        patternCanvas.addEventListener('mousemove', handleCanvasMouseMove);
        
        // Canvas鼠标离开事件
        patternCanvas.addEventListener('mouseleave', handleCanvasMouseLeave);

        // 窗口大小改变事件
        window.addEventListener('resize', debounce(handleResize, 300));
    }

    /**
     * 处理Canvas点击事件
     */
    function handleCanvasClick(event) {
        const rect = patternCanvas.getBoundingClientRect();
        const x = event.clientX - rect.left;
        const y = event.clientY - rect.top;

        // 转换为图案坐标
        const pattern = getPatternAtPosition(x, y);

        if (pattern) {
            // 如果是第一个图案（index 0），直接允许选择
            if (pattern.index === 0) {
                selectPattern(pattern);
                return;
            }

            // 如果还没有检测过黑块，现在检测
            if (!pattern.hasOwnProperty('blackChecked')) {
                const isBlack = isBlackPattern(pattern);
                // 除了第一个黑块（index 0）外，其他黑块都不能选择
                pattern.selectable = !isBlack;
                pattern.blackChecked = true;

                console.log(`图案 ${pattern.index}: 是黑块=${isBlack}, 可选择=${pattern.selectable}`);

                // 重新渲染以显示遮罩
                renderPatternSelector();
            }

            // 检查图案是否可选择
            if (pattern.selectable) {
                selectPattern(pattern);
            } else {
                // 显示不可选择的提示
                showNotSelectableMessage(pattern);
            }
        }
    }

    /**
     * 显示不可选择的提示
     */
    function showNotSelectableMessage(pattern) {
        if (window.App && window.App.PatternSelectorUI && window.App.PatternSelectorUI.showNotification) {
            window.App.PatternSelectorUI.showNotification(
                `图案 ${pattern.index} 不可选择（纯黑块）`,
                'warning',
                2000
            );
        } else {
            // 降级方案：使用原生alert
            console.warn(`图案 ${pattern.index} 不可选择（纯黑块）`);
        }
    }

    /**
     * 处理Canvas鼠标移动事件
     */
    function handleCanvasMouseMove(event) {
        const rect = patternCanvas.getBoundingClientRect();
        const x = event.clientX - rect.left;
        const y = event.clientY - rect.top;

        // 转换为图案坐标
        const pattern = getPatternAtPosition(x, y);
        
        if (pattern) {
            showPatternHover(pattern, x, y);
            patternCanvas.style.cursor = 'pointer';
        } else {
            hidePatternHover();
            patternCanvas.style.cursor = 'default';
        }
    }

    /**
     * 处理Canvas鼠标离开事件
     */
    function handleCanvasMouseLeave() {
        hidePatternHover();
        patternCanvas.style.cursor = 'default';
    }

    /**
     * 根据位置获取图案
     */
    function getPatternAtPosition(x, y) {
        const canvasWidth = patternCanvas.width;
        const canvasHeight = patternCanvas.height;
        
        const cellWidth = canvasWidth / PATTERN_CONFIG.gridCols;
        const cellHeight = canvasHeight / PATTERN_CONFIG.gridRows;
        
        const col = Math.floor(x / cellWidth);
        const row = Math.floor(y / cellHeight);
        
        if (col >= 0 && col < PATTERN_CONFIG.gridCols && 
            row >= 0 && row < PATTERN_CONFIG.gridRows) {
            
            const index = row * PATTERN_CONFIG.gridCols + col;
            return patternGrid[index];
        }
        
        return null;
    }

    /**
     * 选择图案
     */
    function selectPattern(pattern) {
        console.log('选择图案:', pattern);
        
        selectedPattern = pattern;
        
        // 更新视觉反馈
        updatePatternSelection();
        
        // 更新预览
        updatePatternPreview();
        
        // 更新信息显示
        updatePatternInfo();
        
        // 触发选择事件
        triggerPatternSelectEvent(pattern);
    }

    /**
     * 更新图案选择的视觉反馈
     */
    function updatePatternSelection() {
        if (!selectedPattern || !patternCanvas) return;

        // 重新渲染Canvas
        renderPatternSelector();

        // 绘制选中框
        const ctx = patternCanvas.getContext('2d');
        const canvasWidth = patternCanvas.width;
        const canvasHeight = patternCanvas.height;

        const cellWidth = canvasWidth / PATTERN_CONFIG.gridCols;
        const cellHeight = canvasHeight / PATTERN_CONFIG.gridRows;

        const x = selectedPattern.col * cellWidth;
        const y = selectedPattern.row * cellHeight;

        // 绘制选中边框
        ctx.strokeStyle = '#00ff00';
        ctx.lineWidth = 3;
        ctx.strokeRect(x, y, cellWidth, cellHeight);

        // 绘制选中背景
        ctx.fillStyle = 'rgba(0, 255, 0, 0.2)';
        ctx.fillRect(x, y, cellWidth, cellHeight);
    }

    /**
     * 更新图案预览
     */
    function updatePatternPreview() {
        if (!selectedPattern || !patternPreview || !patternImage) return;

        console.log('更新图案预览:', selectedPattern);

        // 创建预览Canvas
        const previewCanvas = document.createElement('canvas');
        const previewSize = 128; // 预览尺寸
        previewCanvas.width = previewSize;
        previewCanvas.height = previewSize;

        const ctx = previewCanvas.getContext('2d');

        // 从原图中提取选中的图案
        ctx.drawImage(
            patternImage,
            selectedPattern.x, selectedPattern.y,
            selectedPattern.width, selectedPattern.height,
            0, 0, previewSize, previewSize
        );

        // 清空预览容器并添加新的预览
        patternPreview.innerHTML = '';
        patternPreview.appendChild(previewCanvas);

        // 添加预览标题
        const previewTitle = document.createElement('div');
        previewTitle.className = 'pattern-preview-title';
        previewTitle.textContent = selectedPattern.name;
        patternPreview.appendChild(previewTitle);
    }

    /**
     * 更新图案信息显示
     */
    function updatePatternInfo() {
        if (!selectedPattern || !patternInfo) return;

        patternInfo.innerHTML = `
            <div class="pattern-info-item">
                <label>图案编号:</label>
                <span>${selectedPattern.name}</span>
            </div>
        `;
    }

    /**
     * 显示图案悬停效果
     */
    function showPatternHover(pattern, x, y) {
        // 显示图案编号和可选择状态
        const status = pattern.selectable ? '可选择' : '不可选择';
        patternCanvas.title = `${pattern.name} (${status})`;
    }

    /**
     * 隐藏图案悬停效果
     */
    function hidePatternHover() {
        patternCanvas.title = '';
    }

    /**
     * 触发图案选择事件
     */
    function triggerPatternSelectEvent(pattern) {
        const event = new CustomEvent('patternSelected', {
            detail: {
                pattern: pattern,
                timestamp: Date.now()
            }
        });

        document.dispatchEvent(event);
        console.log('触发图案选择事件:', pattern);
    }

    /**
     * 处理窗口大小改变
     */
    function handleResize() {
        if (isInitialized && patternImage && patternImage.complete) {
            console.log('窗口大小改变，重新渲染图案选择器');
            renderPatternSelector();

            // 如果有选中的图案，重新绘制选中状态
            if (selectedPattern) {
                updatePatternSelection();
            }
        }
    }

    /**
     * 显示错误信息
     */
    function showError(message) {
        if (patternContainer) {
            patternContainer.innerHTML = `
                <div class="pattern-error">
                    <i class="fas fa-exclamation-triangle"></i>
                    <p>${message}</p>
                </div>
            `;
        }
        console.error('PatternSelector错误:', message);
    }

    /**
     * 重置选择
     */
    function resetSelection() {
        selectedPattern = null;

        if (patternPreview) {
            patternPreview.innerHTML = '<div class="pattern-preview-placeholder">请选择一个图案</div>';
        }

        if (patternInfo) {
            patternInfo.innerHTML = '<div class="pattern-info-placeholder">选择图案后显示详细信息</div>';
        }

        // 重新渲染Canvas（移除选中效果）
        if (isInitialized) {
            renderPatternSelector();
        }
    }

    /**
     * 根据ID选择图案
     */
    function selectPatternById(patternId) {
        const pattern = patternGrid.find(p => p.id === patternId);
        if (pattern) {
            selectPattern(pattern);
            return true;
        }
        return false;
    }

    /**
     * 根据行列选择图案
     */
    function selectPatternByPosition(row, col) {
        if (row >= 0 && row < PATTERN_CONFIG.gridRows &&
            col >= 0 && col < PATTERN_CONFIG.gridCols) {

            const index = row * PATTERN_CONFIG.gridCols + col;
            const pattern = patternGrid[index];
            if (pattern) {
                selectPattern(pattern);
                return true;
            }
        }
        return false;
    }

    // 导出公共接口
    return {
        init: init,
        selectPattern: selectPattern,
        selectPatternById: selectPatternById,
        selectPatternByPosition: selectPatternByPosition,
        resetSelection: resetSelection,
        getSelectedPattern: () => selectedPattern,
        getPatternGrid: () => patternGrid,
        getPatternConfig: () => PATTERN_CONFIG,
        isReady: () => isInitialized
    };

})();

// 工具函数
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

// 模块加载完成日志
console.log('PatternSelector模块已加载 - 图案选择功能已就绪');
