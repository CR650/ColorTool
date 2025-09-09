/**
 * 图案选择器UI控制器
 * 文件说明：处理图案选择器的用户界面交互逻辑
 * 创建时间：2025-09-08
 */

// 确保全局命名空间存在
window.App = window.App || {};

/**
 * 图案选择器UI控制器模块
 */
window.App.PatternSelectorUI = (function() {
    'use strict';

    // 模块状态
    let isInitialized = false;
    let showGrid = true;

    // DOM元素引用
    let resetSelectionBtn = null;
    let toggleGridBtn = null;
    let exportPatternBtn = null;
    let copyPatternBtn = null;
    let usePatternBtn = null;
    let quickSelectBtn = null;
    let rowInput = null;
    let colInput = null;
    let loadingOverlay = null;
    let patternDetailModal = null;
    let closeModalBtn = null;
    let modalCancelBtn = null;
    let modalConfirmBtn = null;

    /**
     * 初始化UI控制器
     */
    function init() {
        if (isInitialized) {
            console.warn('PatternSelectorUI模块已经初始化');
            return;
        }

        console.log('初始化图案选择器UI控制器...');

        // 获取DOM元素
        initializeElements();

        // 绑定事件
        bindEvents();

        // 隐藏加载覆盖层
        hideLoading();

        isInitialized = true;
        console.log('✅ 图案选择器UI控制器初始化完成');
    }

    /**
     * 初始化DOM元素
     */
    function initializeElements() {
        resetSelectionBtn = document.getElementById('resetSelectionBtn');
        toggleGridBtn = document.getElementById('toggleGridBtn');
        exportPatternBtn = document.getElementById('exportPatternBtn');
        copyPatternBtn = document.getElementById('copyPatternBtn');
        usePatternBtn = document.getElementById('usePatternBtn');
        quickSelectBtn = document.getElementById('quickSelectBtn');
        rowInput = document.getElementById('rowInput');
        colInput = document.getElementById('colInput');
        loadingOverlay = document.getElementById('loadingOverlay');
        patternDetailModal = document.getElementById('patternDetailModal');
        closeModalBtn = document.getElementById('closeModalBtn');
        modalCancelBtn = document.getElementById('modalCancelBtn');
        modalConfirmBtn = document.getElementById('modalConfirmBtn');
    }

    /**
     * 绑定事件
     */
    function bindEvents() {
        // 重置选择按钮
        if (resetSelectionBtn) {
            resetSelectionBtn.addEventListener('click', handleResetSelection);
        }

        // 切换网格按钮
        if (toggleGridBtn) {
            toggleGridBtn.addEventListener('click', handleToggleGrid);
        }

        // 导出图案按钮
        if (exportPatternBtn) {
            exportPatternBtn.addEventListener('click', handleExportPattern);
        }

        // 复制图案信息按钮
        if (copyPatternBtn) {
            copyPatternBtn.addEventListener('click', handleCopyPattern);
        }

        // 使用图案按钮
        if (usePatternBtn) {
            usePatternBtn.addEventListener('click', handleUsePattern);
        }

        // 快速选择按钮
        if (quickSelectBtn) {
            quickSelectBtn.addEventListener('click', handleQuickSelect);
        }

        // 输入框回车事件
        if (rowInput) {
            rowInput.addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    handleQuickSelect();
                }
            });
        }

        if (colInput) {
            colInput.addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    handleQuickSelect();
                }
            });
        }

        // 模态框事件
        if (closeModalBtn) {
            closeModalBtn.addEventListener('click', hideModal);
        }

        if (modalCancelBtn) {
            modalCancelBtn.addEventListener('click', hideModal);
        }

        if (modalConfirmBtn) {
            modalConfirmBtn.addEventListener('click', handleModalConfirm);
        }

        // 模态框背景点击关闭
        if (patternDetailModal) {
            patternDetailModal.addEventListener('click', function(e) {
                if (e.target === patternDetailModal) {
                    hideModal();
                }
            });
        }

        // ESC键关闭模态框
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape' && patternDetailModal && patternDetailModal.classList.contains('show')) {
                hideModal();
            }
        });
    }

    /**
     * 处理重置选择
     */
    function handleResetSelection() {
        console.log('重置图案选择');
        
        if (window.App.PatternSelector && window.App.PatternSelector.resetSelection) {
            window.App.PatternSelector.resetSelection();
        }

        // 禁用操作按钮
        disableActionButtons();

        // 清空快速选择输入
        if (rowInput) rowInput.value = '';
        if (colInput) colInput.value = '';

        // 更新状态
        updateStatus('就绪 - 请选择图案');
        updateSelectedCount(0);

        showNotification('已重置选择', 'info');
    }

    /**
     * 处理切换网格显示
     */
    function handleToggleGrid() {
        showGrid = !showGrid;
        console.log('切换网格显示:', showGrid);

        // 更新按钮文本
        if (toggleGridBtn) {
            const icon = toggleGridBtn.querySelector('i');
            if (icon) {
                icon.className = showGrid ? 'fas fa-grid-3x3' : 'fas fa-border-none';
            }
            toggleGridBtn.title = showGrid ? '隐藏网格' : '显示网格';
        }

        // 重新渲染（这里需要PatternSelector模块支持）
        // 可以通过自定义事件通知PatternSelector模块
        const event = new CustomEvent('toggleGrid', {
            detail: { showGrid: showGrid }
        });
        document.dispatchEvent(event);

        showNotification(showGrid ? '已显示网格' : '已隐藏网格', 'info');
    }

    /**
     * 处理导出图案
     */
    function handleExportPattern() {
        console.log('导出选中图案');

        if (!window.App.PatternSelector) {
            showNotification('图案选择器未初始化', 'error');
            return;
        }

        const selectedPattern = window.App.PatternSelector.getSelectedPattern();
        if (!selectedPattern) {
            showNotification('请先选择一个图案', 'warning');
            return;
        }

        try {
            // 创建Canvas来绘制选中的图案
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            
            // 设置Canvas尺寸
            const exportSize = 256; // 导出尺寸
            canvas.width = exportSize;
            canvas.height = exportSize;

            // 获取原始图片
            const patternImage = new Image();
            patternImage.onload = function() {
                // 绘制选中的图案
                ctx.drawImage(
                    this,
                    selectedPattern.x, selectedPattern.y,
                    selectedPattern.width, selectedPattern.height,
                    0, 0, exportSize, exportSize
                );

                // 导出为PNG
                canvas.toBlob(function(blob) {
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `pattern_${selectedPattern.row + 1}_${selectedPattern.col + 1}.png`;
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);

                    showNotification('图案导出成功', 'success');
                }, 'image/png');
            };

            patternImage.src = './Texture/Floor_Pattern.png';

        } catch (error) {
            console.error('导出图案失败:', error);
            showNotification('导出失败: ' + error.message, 'error');
        }
    }

    /**
     * 处理复制图案信息
     */
    function handleCopyPattern() {
        console.log('复制图案信息');

        if (!window.App.PatternSelector) {
            showNotification('图案选择器未初始化', 'error');
            return;
        }

        const selectedPattern = window.App.PatternSelector.getSelectedPattern();
        if (!selectedPattern) {
            showNotification('请先选择一个图案', 'warning');
            return;
        }

        const patternInfo = {
            name: selectedPattern.name,
            id: selectedPattern.id,
            position: `行 ${selectedPattern.row + 1}, 列 ${selectedPattern.col + 1}`,
            coordinates: `(${selectedPattern.x}, ${selectedPattern.y})`,
            size: `${selectedPattern.width} x ${selectedPattern.height}`
        };

        const infoText = Object.entries(patternInfo)
            .map(([key, value]) => `${key}: ${value}`)
            .join('\n');

        // 复制到剪贴板
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(infoText).then(() => {
                showNotification('图案信息已复制到剪贴板', 'success');
            }).catch(err => {
                console.error('复制失败:', err);
                showNotification('复制失败', 'error');
            });
        } else {
            // 降级方案
            const textArea = document.createElement('textarea');
            textArea.value = infoText;
            document.body.appendChild(textArea);
            textArea.select();
            try {
                document.execCommand('copy');
                showNotification('图案信息已复制到剪贴板', 'success');
            } catch (err) {
                console.error('复制失败:', err);
                showNotification('复制失败', 'error');
            }
            document.body.removeChild(textArea);
        }
    }

    /**
     * 处理使用图案
     */
    function handleUsePattern() {
        console.log('使用选中图案');

        if (!window.App.PatternSelector) {
            showNotification('图案选择器未初始化', 'error');
            return;
        }

        const selectedPattern = window.App.PatternSelector.getSelectedPattern();
        if (!selectedPattern) {
            showNotification('请先选择一个图案', 'warning');
            return;
        }

        // 显示确认模态框
        showPatternDetailModal(selectedPattern);
    }

    /**
     * 处理快速选择
     */
    function handleQuickSelect() {
        const row = parseInt(rowInput?.value) - 1; // 转换为0基索引
        const col = parseInt(colInput?.value) - 1; // 转换为0基索引

        if (isNaN(row) || isNaN(col)) {
            showNotification('请输入有效的行列数字', 'warning');
            return;
        }

        if (!window.App.PatternSelector) {
            showNotification('图案选择器未初始化', 'error');
            return;
        }

        const success = window.App.PatternSelector.selectPatternByPosition(row, col);
        if (success) {
            showNotification(`已选择图案 (${row + 1}, ${col + 1})`, 'success');
        } else {
            showNotification('选择失败，请检查行列范围', 'error');
        }
    }

    /**
     * 显示加载状态
     */
    function showLoading() {
        if (loadingOverlay) {
            loadingOverlay.style.display = 'flex';
        }
    }

    /**
     * 隐藏加载状态
     */
    function hideLoading() {
        if (loadingOverlay) {
            loadingOverlay.style.display = 'none';
        }
    }

    /**
     * 显示通知消息
     */
    function showNotification(message, type = 'info', duration = 3000) {
        // 创建通知元素
        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;
        notification.innerHTML = `
            <i class="fas fa-${getNotificationIcon(type)}"></i>
            <span>${message}</span>
        `;

        // 添加样式
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: ${getNotificationColor(type)};
            color: white;
            padding: 12px 20px;
            border-radius: 8px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
            z-index: 1001;
            display: flex;
            align-items: center;
            gap: 8px;
            font-weight: 500;
            transform: translateX(100%);
            transition: transform 0.3s ease;
        `;

        document.body.appendChild(notification);

        // 显示动画
        setTimeout(() => {
            notification.style.transform = 'translateX(0)';
        }, 100);

        // 自动隐藏
        setTimeout(() => {
            notification.style.transform = 'translateX(100%)';
            setTimeout(() => {
                if (notification.parentNode) {
                    notification.parentNode.removeChild(notification);
                }
            }, 300);
        }, duration);
    }

    /**
     * 获取通知图标
     */
    function getNotificationIcon(type) {
        const icons = {
            success: 'check-circle',
            error: 'exclamation-circle',
            warning: 'exclamation-triangle',
            info: 'info-circle'
        };
        return icons[type] || 'info-circle';
    }

    /**
     * 获取通知颜色
     */
    function getNotificationColor(type) {
        const colors = {
            success: '#28a745',
            error: '#dc3545',
            warning: '#ffc107',
            info: '#17a2b8'
        };
        return colors[type] || '#17a2b8';
    }

    /**
     * 更新状态文本
     */
    function updateStatus(text) {
        const statusText = document.getElementById('statusText');
        if (statusText) {
            statusText.textContent = text;
        }
    }

    /**
     * 更新选择计数
     */
    function updateSelectedCount(count) {
        const selectedCount = document.getElementById('selectedCount');
        if (selectedCount) {
            selectedCount.textContent = count.toString();
        }
    }

    /**
     * 禁用操作按钮
     */
    function disableActionButtons() {
        const buttons = [exportPatternBtn, copyPatternBtn, usePatternBtn];
        buttons.forEach(btn => {
            if (btn) btn.disabled = true;
        });
    }

    /**
     * 启用操作按钮
     */
    function enableActionButtons() {
        const buttons = [exportPatternBtn, copyPatternBtn, usePatternBtn];
        buttons.forEach(btn => {
            if (btn) btn.disabled = false;
        });
    }

    /**
     * 显示图案详情模态框
     */
    function showPatternDetailModal(pattern) {
        if (!patternDetailModal) return;

        // 填充模态框内容
        const modalContent = document.getElementById('patternDetailContent');
        if (modalContent) {
            modalContent.innerHTML = `
                <div class="pattern-detail">
                    <div class="pattern-detail-preview">
                        <canvas id="modalPatternCanvas" width="200" height="200"></canvas>
                    </div>
                    <div class="pattern-detail-info">
                        <h4>${pattern.name}</h4>
                        <p><strong>状态:</strong> ${pattern.selectable ? '可选择' : '不可选择'}</p>
                    </div>
                </div>
            `;

            // 绘制图案预览
            const canvas = document.getElementById('modalPatternCanvas');
            if (canvas) {
                const ctx = canvas.getContext('2d');
                const img = new Image();
                img.onload = function() {
                    ctx.drawImage(
                        this,
                        pattern.x, pattern.y,
                        pattern.width, pattern.height,
                        0, 0, 200, 200
                    );
                };
                img.src = './Texture/Floor_Pattern.png';
            }
        }

        // 显示模态框
        patternDetailModal.classList.add('show');
    }

    /**
     * 隐藏模态框
     */
    function hideModal() {
        if (patternDetailModal) {
            patternDetailModal.classList.remove('show');
        }
    }

    /**
     * 处理模态框确认
     */
    function handleModalConfirm() {
        const selectedPattern = window.App.PatternSelector?.getSelectedPattern();
        if (selectedPattern) {
            // 触发使用图案事件
            const event = new CustomEvent('patternUsed', {
                detail: {
                    pattern: selectedPattern,
                    timestamp: Date.now()
                }
            });
            document.dispatchEvent(event);

            showNotification(`已使用图案: ${selectedPattern.name}`, 'success');
            hideModal();
        }
    }

    // 导出公共接口
    return {
        init: init,
        showLoading: showLoading,
        hideLoading: hideLoading,
        showNotification: showNotification,
        updateStatus: updateStatus,
        updateSelectedCount: updateSelectedCount,
        enableActionButtons: enableActionButtons,
        disableActionButtons: disableActionButtons,
        showPatternDetailModal: showPatternDetailModal,
        hideModal: hideModal,
        isReady: () => isInitialized
    };

})();

// 模块加载完成日志
console.log('PatternSelectorUI模块已加载 - 图案选择器UI控制功能已就绪');
