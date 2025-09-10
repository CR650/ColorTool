/**
 * 颜色选择器模块
 * 文件说明：为Light配置提供颜色选择功能
 * 创建时间：2025-01-09
 */

// 确保全局命名空间存在
window.App = window.App || {};

/**
 * 颜色选择器模块
 */
window.App.ColorPicker = (function() {
    'use strict';

    // 模块状态
    let isInitialized = false;
    let currentTargetInput = null;
    let currentMode = 'hex'; // 'hex' 或 'rgb'
    let currentRgbTarget = null; // RGB模式下的目标前缀

    // DOM元素引用
    let modal = null;
    let colorCanvas = null;
    let ctx = null;
    let confirmBtn = null;
    let cancelBtn = null;
    let closeBtn = null;
    let hexInput = null;
    let colorPreview = null;
    let selectedColor = 'FFFFFF';

    /**
     * 初始化颜色选择器
     */
    function init() {
        if (isInitialized) {
            console.warn('ColorPicker模块已经初始化');
            return;
        }

        console.log('初始化颜色选择器模块...');

        // 创建颜色选择器模态框
        createColorPickerModal();

        // 绑定事件
        bindEvents();

        isInitialized = true;
        console.log('颜色选择器模块初始化完成');
    }

    /**
     * 创建颜色选择器模态框
     */
    function createColorPickerModal() {
        // 创建模态框HTML
        const modalHTML = `
            <div id="colorPickerModal" class="color-picker-modal" style="display: none;">
                <div class="color-picker-modal-content">
                    <div class="color-picker-modal-header">
                        <h3>选择颜色</h3>
                        <button class="color-picker-modal-close" id="closeColorPickerModal">&times;</button>
                    </div>
                    <div class="color-picker-modal-body">
                        <div class="color-picker-container">
                            <canvas id="colorPickerCanvas" class="color-picker-canvas" width="300" height="200"></canvas>
                        </div>
                        <div class="color-input-container">
                            <label for="colorHexInput">16进制颜色值:</label>
                            <input type="text" id="colorHexInput" maxlength="6" placeholder="FFFFFF" pattern="[0-9A-Fa-f]{6}">
                            <div class="color-preview-large" id="colorPreviewLarge"></div>
                        </div>
                        <div class="preset-colors">
                            <h4>常用颜色</h4>
                            <div class="preset-color-grid">
                                <div class="preset-color" data-color="FFFFFF" style="background-color: #FFFFFF;" title="白色"></div>
                                <div class="preset-color" data-color="000000" style="background-color: #000000;" title="黑色"></div>
                                <div class="preset-color" data-color="FF0000" style="background-color: #FF0000;" title="红色"></div>
                                <div class="preset-color" data-color="00FF00" style="background-color: #00FF00;" title="绿色"></div>
                                <div class="preset-color" data-color="0000FF" style="background-color: #0000FF;" title="蓝色"></div>
                                <div class="preset-color" data-color="FFFF00" style="background-color: #FFFF00;" title="黄色"></div>
                                <div class="preset-color" data-color="FF00FF" style="background-color: #FF00FF;" title="洋红"></div>
                                <div class="preset-color" data-color="00FFFF" style="background-color: #00FFFF;" title="青色"></div>
                                <div class="preset-color" data-color="808080" style="background-color: #808080;" title="灰色"></div>
                                <div class="preset-color" data-color="C0C0C0" style="background-color: #C0C0C0;" title="银色"></div>
                                <div class="preset-color" data-color="800000" style="background-color: #800000;" title="栗色"></div>
                                <div class="preset-color" data-color="008000" style="background-color: #008000;" title="深绿"></div>
                                <div class="preset-color" data-color="000080" style="background-color: #000080;" title="深蓝"></div>
                                <div class="preset-color" data-color="808000" style="background-color: #808000;" title="橄榄"></div>
                                <div class="preset-color" data-color="800080" style="background-color: #800080;" title="紫色"></div>
                                <div class="preset-color" data-color="008080" style="background-color: #008080;" title="青绿"></div>
                            </div>
                        </div>
                    </div>
                    <div class="color-picker-modal-footer">
                        <button id="confirmColorSelection" class="color-picker-btn color-picker-btn-primary">确认选择</button>
                        <button id="cancelColorSelection" class="color-picker-btn color-picker-btn-secondary">取消</button>
                    </div>
                </div>
            </div>
        `;

        // 添加到页面
        document.body.insertAdjacentHTML('beforeend', modalHTML);

        // 获取DOM元素引用
        modal = document.getElementById('colorPickerModal');
        colorCanvas = document.getElementById('colorPickerCanvas');
        ctx = colorCanvas.getContext('2d');
        confirmBtn = document.getElementById('confirmColorSelection');
        cancelBtn = document.getElementById('cancelColorSelection');
        closeBtn = document.getElementById('closeColorPickerModal');
        hexInput = document.getElementById('colorHexInput');
        colorPreview = document.getElementById('colorPreviewLarge');

        // 绘制颜色选择器
        drawColorPicker();
    }

    /**
     * 绘制颜色选择器
     */
    function drawColorPicker() {
        const width = colorCanvas.width;
        const height = colorCanvas.height;

        // 创建渐变
        for (let x = 0; x < width; x++) {
            for (let y = 0; y < height; y++) {
                const hue = (x / width) * 360;
                const saturation = 100;
                const lightness = ((height - y) / height) * 100;
                
                const color = hslToRgb(hue, saturation, lightness);
                const hex = rgbToHex(color.r, color.g, color.b);
                
                ctx.fillStyle = `#${hex}`;
                ctx.fillRect(x, y, 1, 1);
            }
        }
    }

    /**
     * HSL转RGB
     */
    function hslToRgb(h, s, l) {
        h /= 360;
        s /= 100;
        l /= 100;

        const c = (1 - Math.abs(2 * l - 1)) * s;
        const x = c * (1 - Math.abs((h * 6) % 2 - 1));
        const m = l - c / 2;

        let r, g, b;

        if (0 <= h && h < 1/6) {
            r = c; g = x; b = 0;
        } else if (1/6 <= h && h < 2/6) {
            r = x; g = c; b = 0;
        } else if (2/6 <= h && h < 3/6) {
            r = 0; g = c; b = x;
        } else if (3/6 <= h && h < 4/6) {
            r = 0; g = x; b = c;
        } else if (4/6 <= h && h < 5/6) {
            r = x; g = 0; b = c;
        } else {
            r = c; g = 0; b = x;
        }

        return {
            r: Math.round((r + m) * 255),
            g: Math.round((g + m) * 255),
            b: Math.round((b + m) * 255)
        };
    }

    /**
     * RGB转16进制
     */
    function rgbToHex(r, g, b) {
        return ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
    }

    /**
     * 绑定事件监听器
     */
    function bindEvents() {
        // 绑定颜色选择按钮
        document.addEventListener('click', function(e) {
            if (e.target.classList.contains('color-picker-btn') && e.target.hasAttribute('data-target')) {
                const targetId = e.target.getAttribute('data-target');
                const mode = e.target.getAttribute('data-mode') || 'hex';
                openColorPicker(targetId, mode);
            }
        });

        // 模态框关闭事件
        if (closeBtn) {
            closeBtn.addEventListener('click', closeModal);
        }

        if (cancelBtn) {
            cancelBtn.addEventListener('click', closeModal);
        }

        if (confirmBtn) {
            confirmBtn.addEventListener('click', confirmSelection);
        }

        // 画布点击事件
        if (colorCanvas) {
            colorCanvas.addEventListener('click', handleCanvasClick);
        }

        // 16进制输入事件
        if (hexInput) {
            hexInput.addEventListener('input', handleHexInput);
        }

        // 预设颜色点击事件
        document.addEventListener('click', function(e) {
            if (e.target.classList.contains('preset-color')) {
                const color = e.target.getAttribute('data-color');
                selectColor(color);
            }
        });

        // 颜色预览更新事件
        document.addEventListener('input', function(e) {
            if (e.target.id === 'lightSpecularColor') {
                updateColorPreview('lightSpecularColor', e.target.value);
            }
        });
    }

    /**
     * 打开颜色选择器
     * @param {string} targetId - 目标输入框ID或RGB组名
     * @param {string} mode - 模式：'hex' 或 'rgb'
     */
    function openColorPicker(targetId, mode = 'hex') {
        currentMode = mode;

        if (mode === 'rgb') {
            // RGB模式：targetId是RGB组的前缀（如 'pickupDiff'）
            currentRgbTarget = targetId;
            const rInput = document.getElementById(targetId + 'R');
            const gInput = document.getElementById(targetId + 'G');
            const bInput = document.getElementById(targetId + 'B');

            if (!rInput || !gInput || !bInput) {
                console.error('RGB输入框未找到:', targetId);
                return;
            }

            // 从RGB值获取当前颜色
            const r = parseInt(rInput.value || '255');
            const g = parseInt(gInput.value || '255');
            const b = parseInt(bInput.value || '255');
            const hexColor = rgbToHex(r, g, b);
            selectColor(hexColor);

        } else {
            // HEX模式：targetId是输入框ID
            currentTargetInput = document.getElementById(targetId);
            currentRgbTarget = null;

            if (!currentTargetInput) {
                console.error('目标输入框未找到:', targetId);
                return;
            }

            // 获取当前颜色值
            const currentValue = currentTargetInput.value || 'FFFFFF';
            selectColor(currentValue);
        }

        // 显示模态框
        modal.style.display = 'flex';

        console.log('颜色选择器已打开，目标:', targetId, '模式:', mode);
    }

    /**
     * RGB转16进制
     */
    function rgbToHex(r, g, b) {
        const toHex = (n) => {
            const hex = Math.max(0, Math.min(255, Math.round(n))).toString(16);
            return hex.length === 1 ? '0' + hex : hex;
        };
        return (toHex(r) + toHex(g) + toHex(b)).toUpperCase();
    }

    /**
     * 16进制转RGB
     */
    function hexToRgb(hex) {
        const cleanHex = hex.replace('#', '').toUpperCase();
        if (cleanHex.length !== 6) return { r: 255, g: 255, b: 255 };

        const r = parseInt(cleanHex.substr(0, 2), 16);
        const g = parseInt(cleanHex.substr(2, 2), 16);
        const b = parseInt(cleanHex.substr(4, 2), 16);

        return { r, g, b };
    }

    /**
     * 关闭模态框
     */
    function closeModal() {
        modal.style.display = 'none';
        currentTargetInput = null;
        currentRgbTarget = null;
        currentMode = 'hex';
    }

    /**
     * 确认选择
     */
    function confirmSelection() {
        if (currentMode === 'rgb' && currentRgbTarget) {
            // RGB模式：将颜色拆分为R、G、B值
            const rgb = hexToRgb(selectedColor);
            const rInput = document.getElementById(currentRgbTarget + 'R');
            const gInput = document.getElementById(currentRgbTarget + 'G');
            const bInput = document.getElementById(currentRgbTarget + 'B');
            const preview = document.getElementById(currentRgbTarget + 'ColorPreview');

            if (rInput && gInput && bInput) {
                rInput.value = rgb.r;
                gInput.value = rgb.g;
                bInput.value = rgb.b;

                // 更新颜色预览
                if (preview) {
                    preview.style.backgroundColor = `rgb(${rgb.r}, ${rgb.g}, ${rgb.b})`;
                }

                // 触发验证事件
                [rInput, gInput, bInput].forEach(input => {
                    input.dispatchEvent(new Event('input', { bubbles: true }));
                });

                console.log(`RGB颜色已设置: ${currentRgbTarget} = RGB(${rgb.r}, ${rgb.g}, ${rgb.b})`);
            }

        } else if (currentTargetInput) {
            // HEX模式：直接设置16进制值
            currentTargetInput.value = selectedColor;
            updateColorPreview(currentTargetInput.id, selectedColor);
        }

        closeModal();
    }

    /**
     * 处理画布点击
     */
    function handleCanvasClick(e) {
        const rect = colorCanvas.getBoundingClientRect();
        const x = e.clientX - rect.left;
        const y = e.clientY - rect.top;

        const imageData = ctx.getImageData(x, y, 1, 1);
        const data = imageData.data;
        
        const hex = rgbToHex(data[0], data[1], data[2]);
        selectColor(hex);
    }

    /**
     * 处理16进制输入
     */
    function handleHexInput(e) {
        const value = e.target.value.toUpperCase();
        if (/^[0-9A-F]{0,6}$/i.test(value)) {
            if (value.length === 6) {
                selectColor(value);
            }
        }
    }

    /**
     * 选择颜色
     */
    function selectColor(color) {
        selectedColor = color.replace('#', '').toUpperCase();
        
        // 更新输入框
        if (hexInput) {
            hexInput.value = selectedColor;
        }
        
        // 更新预览
        if (colorPreview) {
            colorPreview.style.backgroundColor = `#${selectedColor}`;
        }
    }

    /**
     * 更新颜色预览
     */
    function updateColorPreview(inputId, color) {
        const previewId = inputId + 'Preview';
        const preview = document.getElementById(previewId);
        
        if (preview) {
            const cleanColor = color.replace('#', '');
            preview.style.backgroundColor = `#${cleanColor}`;
        }
    }

    // 返回公共接口
    return {
        init: init,
        updateColorPreview: updateColorPreview
    };

})();
