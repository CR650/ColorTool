/**
 * UGC图案选择器模块
 * 文件说明：为UGC配置提供图案选择功能
 * 创建时间：2025-01-09
 */

// 确保全局命名空间存在
window.App = window.App || {};

/**
 * UGC图案选择器模块
 */
window.App.UGCPatternSelector = (function() {
    'use strict';

    // 模块状态
    let isInitialized = false;
    let currentTargetInput = null;
    let currentImageType = null;
    let patternSelector = null;

    // DOM元素引用
    let modal = null;
    let modalTitle = null;
    let canvas = null;
    let ctx = null;
    let loadingText = null;
    let selectedPatternInfo = null;
    let patternPreviewCanvas = null;
    let confirmBtn = null;
    let cancelBtn = null;
    let closeBtn = null;

    // 图案配置
    const PATTERN_CONFIG = {
        gridCols: 8,
        gridRows: 8,
        images: {
            pattern: './Texture/Floor_Pattern.png',
            frame: './Texture/Floor_Board.png'
        }
    };

    /**
     * 初始化UGC图案选择器
     */
    function init() {
        if (isInitialized) {
            console.warn('UGCPatternSelector模块已经初始化');
            return;
        }

        console.log('初始化UGC图案选择器模块...');

        // 获取DOM元素
        modal = document.getElementById('patternSelectorModal');
        modalTitle = document.getElementById('patternModalTitle');
        canvas = document.getElementById('patternSelectorCanvas');
        ctx = canvas.getContext('2d');
        loadingText = document.getElementById('patternLoadingText');
        selectedPatternInfo = document.getElementById('selectedPatternInfo');
        patternPreviewCanvas = document.getElementById('patternPreviewCanvas');
        confirmBtn = document.getElementById('confirmPatternSelection');
        cancelBtn = document.getElementById('cancelPatternSelection');
        closeBtn = document.getElementById('closePatternModal');

        if (!modal || !canvas) {
            console.error('UGC图案选择器必需的DOM元素未找到');
            return;
        }

        // 绑定事件
        bindEvents();

        isInitialized = true;
        console.log('UGC图案选择器模块初始化完成');
    }

    /**
     * 绑定事件监听器
     */
    function bindEvents() {
        // 绑定所有图案选择按钮 - 使用事件委托
        document.addEventListener('click', function(e) {
            if (e.target.classList.contains('pattern-select-btn')) {
                const targetId = e.target.getAttribute('data-target');
                const imageType = e.target.getAttribute('data-type');

                e.preventDefault();
                e.stopPropagation();

                openPatternSelector(targetId, imageType);
            }
        }, true); // 使用捕获阶段

        // 模态框关闭事件
        closeBtn.addEventListener('click', closeModal);
        cancelBtn.addEventListener('click', closeModal);
        
        // 点击模态框背景关闭
        modal.addEventListener('click', function(e) {
            if (e.target === modal) {
                closeModal();
            }
        });

        // 确认选择
        confirmBtn.addEventListener('click', confirmSelection);

        // 画布点击事件
        canvas.addEventListener('click', handleCanvasClick);
    }

    /**
     * 打开图案选择器
     */
    function openPatternSelector(targetInputId, imageType) {
        currentTargetInput = document.getElementById(targetInputId);
        currentImageType = imageType;

        if (!currentTargetInput) {
            console.error('目标输入框未找到:', targetInputId);
            return;
        }

        // 设置模态框标题，包含当前值
        const currentValue = currentTargetInput.value || '0';
        const titles = {
            pattern: `选择图案 (Floor_Pattern) - 当前: ${currentValue}`,
            frame: `选择边框 (Floor_Board) - 当前: ${currentValue}`
        };
        modalTitle.textContent = titles[imageType] || `选择图案 - 当前: ${currentValue}`;

        // 显示模态框
        modal.style.display = 'flex';
        
        // 重置状态
        selectedPatternInfo.style.display = 'none';
        confirmBtn.disabled = true;
        loadingText.style.display = 'block';

        // 加载图案
        loadPatternImage(imageType);
    }

    /**
     * 加载图案图片
     */
    function loadPatternImage(imageType) {
        const imagePath = PATTERN_CONFIG.images[imageType];
        
        if (!imagePath) {
            console.error('未知的图片类型:', imageType);
            loadingText.textContent = '图片类型错误';
            return;
        }

        const image = new Image();
        image.onload = function() {
            setupCanvas(image);
            createPatternSelector(image);

            // 预选当前输入框中的图案ID
            preselectCurrentPattern();

            loadingText.style.display = 'none';
        };
        
        image.onerror = function() {
            loadingText.textContent = '图片加载失败，请检查路径: ' + imagePath;
            console.error('图片加载失败:', imagePath);
        };
        
        image.src = imagePath;
    }

    /**
     * 设置画布尺寸
     */
    function setupCanvas(image) {
        const containerWidth = canvas.parentElement.clientWidth - 40;
        const aspectRatio = image.height / image.width;
        const canvasWidth = Math.min(containerWidth, 600);
        const canvasHeight = canvasWidth * aspectRatio;
        
        canvas.width = canvasWidth;
        canvas.height = canvasHeight;
        canvas.style.width = canvasWidth + 'px';
        canvas.style.height = canvasHeight + 'px';
    }

    /**
     * 创建图案选择器实例
     */
    function createPatternSelector(image) {
        patternSelector = new PatternGrid(image, canvas, PATTERN_CONFIG.gridCols, PATTERN_CONFIG.gridRows);
        patternSelector.render();
    }

    /**
     * 预选当前输入框中的图案ID
     */
    function preselectCurrentPattern() {
        if (!patternSelector || !currentTargetInput) {
            return;
        }

        const currentValue = parseInt(currentTargetInput.value || '0');
        console.log(`预选图案ID: ${currentValue}`);

        // 查找对应的图案
        const targetPattern = patternSelector.patterns.find(pattern => pattern.index === currentValue);

        if (targetPattern) {
            console.log(`找到对应图案:`, targetPattern);
            patternSelector.selectPattern(targetPattern);
            showSelectedPattern(targetPattern);
            confirmBtn.disabled = false;
        } else {
            console.log(`未找到图案ID ${currentValue}，使用默认选择`);
            // 如果没找到，默认选择第一个图案（ID=0）
            if (patternSelector.patterns.length > 0) {
                const defaultPattern = patternSelector.patterns[0];
                patternSelector.selectPattern(defaultPattern);
                showSelectedPattern(defaultPattern);
                confirmBtn.disabled = false;
            }
        }
    }

    /**
     * 处理画布点击事件
     */
    function handleCanvasClick(event) {
        if (!patternSelector) return;

        const rect = canvas.getBoundingClientRect();
        const x = event.clientX - rect.left;
        const y = event.clientY - rect.top;

        const selectedPattern = patternSelector.getPatternAt(x, y);
        if (selectedPattern) {
            patternSelector.selectPattern(selectedPattern);
            showSelectedPattern(selectedPattern);
            confirmBtn.disabled = false;
        }
    }

    /**
     * 显示选中的图案信息
     */
    function showSelectedPattern(pattern) {
        selectedPatternInfo.style.display = 'block';
        
        // 更新图案信息
        document.getElementById('selectedPatternId').textContent = pattern.index;
        document.getElementById('selectedPatternPosition').textContent = `行 ${pattern.row + 1}, 列 ${pattern.col + 1}`;
        
        // 绘制预览
        drawPatternPreview(pattern);
    }

    /**
     * 绘制图案预览
     */
    function drawPatternPreview(pattern) {
        patternPreviewCanvas.width = 100;
        patternPreviewCanvas.height = 100;
        
        const previewCtx = patternPreviewCanvas.getContext('2d');
        previewCtx.drawImage(
            patternSelector.image,
            pattern.x, pattern.y, pattern.width, pattern.height,
            0, 0, 100, 100
        );
    }

    /**
     * 确认选择
     */
    function confirmSelection() {
        if (!patternSelector || !patternSelector.selectedPattern || !currentTargetInput) {
            return;
        }

        const selectedIndex = patternSelector.selectedPattern.index;
        currentTargetInput.value = selectedIndex;
        
        // 触发input事件以便其他模块监听
        currentTargetInput.dispatchEvent(new Event('input', { bubbles: true }));
        
        closeModal();
    }

    /**
     * 关闭模态框
     */
    function closeModal() {
        modal.style.display = 'none';
        currentTargetInput = null;
        currentImageType = null;
        patternSelector = null;
        
        // 清空画布
        if (ctx) {
            ctx.clearRect(0, 0, canvas.width, canvas.height);
        }
    }

    /**
     * 图案网格类
     */
    class PatternGrid {
        constructor(image, canvas, cols, rows) {
            this.image = image;
            this.canvas = canvas;
            this.ctx = canvas.getContext('2d');
            this.cols = cols;
            this.rows = rows;
            this.patterns = [];
            this.selectedPattern = null;
            
            this.generatePatterns();
        }

        generatePatterns() {
            this.patterns = [];
            const patternWidth = this.image.width / this.cols;
            const patternHeight = this.image.height / this.rows;

            let index = 0;
            for (let row = 0; row < this.rows; row++) {
                for (let col = 0; col < this.cols; col++) {
                    this.patterns.push({
                        index: index,
                        row: row,
                        col: col,
                        x: col * patternWidth,
                        y: row * patternHeight,
                        width: patternWidth,
                        height: patternHeight
                    });
                    index++;
                }
            }
        }

        render() {
            // 绘制图片
            this.ctx.drawImage(this.image, 0, 0, this.canvas.width, this.canvas.height);
            
            // 绘制网格线
            this.drawGrid();
            
            // 绘制选中状态
            if (this.selectedPattern) {
                this.drawSelection();
            }
        }

        drawGrid() {
            const cellWidth = this.canvas.width / this.cols;
            const cellHeight = this.canvas.height / this.rows;
            
            this.ctx.strokeStyle = 'rgba(255, 255, 255, 0.5)';
            this.ctx.lineWidth = 1;
            
            // 绘制垂直线
            for (let i = 1; i < this.cols; i++) {
                const x = i * cellWidth;
                this.ctx.beginPath();
                this.ctx.moveTo(x, 0);
                this.ctx.lineTo(x, this.canvas.height);
                this.ctx.stroke();
            }
            
            // 绘制水平线
            for (let i = 1; i < this.rows; i++) {
                const y = i * cellHeight;
                this.ctx.beginPath();
                this.ctx.moveTo(0, y);
                this.ctx.lineTo(this.canvas.width, y);
                this.ctx.stroke();
            }
        }

        drawSelection() {
            const cellWidth = this.canvas.width / this.cols;
            const cellHeight = this.canvas.height / this.rows;
            const x = this.selectedPattern.col * cellWidth;
            const y = this.selectedPattern.row * cellHeight;

            // 绘制选中边框（更粗的边框）
            this.ctx.strokeStyle = '#00ff00';
            this.ctx.lineWidth = 4;
            this.ctx.strokeRect(x + 2, y + 2, cellWidth - 4, cellHeight - 4);

            // 绘制半透明覆盖
            this.ctx.fillStyle = 'rgba(0, 255, 0, 0.3)';
            this.ctx.fillRect(x, y, cellWidth, cellHeight);

            // 绘制图案ID标识
            this.ctx.fillStyle = '#00ff00';
            this.ctx.font = 'bold 16px Arial';
            this.ctx.textAlign = 'center';
            this.ctx.textBaseline = 'middle';

            // 绘制白色背景
            const textX = x + cellWidth / 2;
            const textY = y + cellHeight / 2;
            const textWidth = 30;
            const textHeight = 20;

            this.ctx.fillStyle = 'rgba(255, 255, 255, 0.9)';
            this.ctx.fillRect(textX - textWidth/2, textY - textHeight/2, textWidth, textHeight);

            // 绘制图案ID文字
            this.ctx.fillStyle = '#000000';
            this.ctx.fillText(this.selectedPattern.index.toString(), textX, textY);
        }

        getPatternAt(x, y) {
            const cellWidth = this.canvas.width / this.cols;
            const cellHeight = this.canvas.height / this.rows;
            
            const col = Math.floor(x / cellWidth);
            const row = Math.floor(y / cellHeight);
            
            if (col >= 0 && col < this.cols && row >= 0 && row < this.rows) {
                const index = row * this.cols + col;
                return this.patterns[index];
            }
            
            return null;
        }

        selectPattern(pattern) {
            this.selectedPattern = pattern;
            this.render();
        }
    }

    /**
     * 重新绑定按钮事件（当UGC面板显示时调用）
     */
    function rebindButtonEvents() {
        const buttons = document.querySelectorAll('.pattern-select-btn');

        buttons.forEach((button) => {
            // 移除之前的监听器（如果有）
            button.removeEventListener('click', handleButtonClick);

            // 添加新的监听器
            button.addEventListener('click', handleButtonClick);
        });
    }

    /**
     * 按钮点击处理函数
     */
    function handleButtonClick(e) {
        e.preventDefault();
        e.stopPropagation();

        const targetId = e.target.getAttribute('data-target');
        const imageType = e.target.getAttribute('data-type');

        openPatternSelector(targetId, imageType);
    }

    // 暴露公共接口
    return {
        init: init,
        rebindButtonEvents: rebindButtonEvents
    };

})();

// 模块加载完成日志
console.log('UGCPatternSelector模块已加载');
