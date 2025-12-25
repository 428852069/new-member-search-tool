// == 新发展团员编号批量搜索工具（CDN版）==
// GitHub托管版本
// 使用方法：在新发展团员搜索页面运行此脚本

(function() {
    'use strict';
    
    // 检查是否已加载
    if (window.newMemberSearchLoaded) {
        alert('工具已加载，无需重复运行');
        return;
    }
    window.newMemberSearchLoaded = true;
    
    // 配置参数
    const config = {
        delayBetweenSearches: 2000, // 每次搜索间隔（毫秒）
        resultWaitTime: 3000,       // 等待结果的最大时间（毫秒）
        saveAsCSV: true            // 是否保存为CSV格式
    };
    
    // 结果存储
    let allResults = [];
    let searchIndex = 0;
    let isRunning = false;
    let searchItems = [];
    
    // 创建输入框让用户粘贴数据
    function createDataInputModal() {
        const modal = document.createElement('div');
        modal.id = 'dataInputModal';
        modal.style.cssText = `
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: white;
            border: 2px solid #409EFF;
            border-radius: 8px;
            padding: 20px;
            z-index: 999999;
            box-shadow: 0 4px 20px rgba(0,0,0,0.2);
            width: 600px;
            max-width: 90vw;
            max-height: 80vh;
            overflow: hidden;
            font-family: Arial, sans-serif;
        `;
        
        modal.innerHTML = `
            <div style="margin-bottom: 15px; font-weight: bold; color: #409EFF; font-size: 18px;">
                新发展团员编号批量搜索工具
            </div>
            <div style="margin-bottom: 10px; color: #666;">
                请在下方粘贴Excel中的新发展团员编号（每行一个）：
            </div>
            <textarea 
                id="searchDataInput" 
                style="
                    width: 100%;
                    height: 300px;
                    padding: 10px;
                    border: 1px solid #dcdfe6;
                    border-radius: 4px;
                    font-family: 'Consolas', 'Monaco', monospace;
                    font-size: 14px;
                    resize: vertical;
                    margin-bottom: 15px;
                "
                placeholder="示例：
20210001
20210002
20210003
..."
            ></textarea>
            <div style="color: #999; font-size: 12px; margin-bottom: 15px;">
                提示：直接从Excel中复制新发展团员编号，然后粘贴到上面的文本框
            </div>
            <div style="display: flex; gap: 10px; justify-content: flex-end;">
                <button 
                    id="cancelBtn" 
                    style="padding: 8px 20px; background: #909399; color: white; border: none; border-radius: 4px; cursor: pointer;"
                >
                    取消
                </button>
                <button 
                    id="startBtn" 
                    style="padding: 8px 20px; background: #409EFF; color: white; border: none; border-radius: 4px; cursor: pointer;"
                >
                    开始搜索
                </button>
            </div>
        `;
        
        document.body.appendChild(modal);
        
        // 添加遮罩层
        const overlay = document.createElement('div');
        overlay.id = 'modalOverlay';
        overlay.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.5);
            z-index: 999998;
        `;
        document.body.appendChild(overlay);
        
        // 焦点自动到输入框
        setTimeout(() => {
            document.getElementById('searchDataInput').focus();
        }, 100);
        
        // 绑定事件
        document.getElementById('startBtn').addEventListener('click', startFromInput);
        document.getElementById('cancelBtn').addEventListener('click', () => {
            modal.remove();
            overlay.remove();
            alert('已取消批量搜索');
        });
        
        // ESC键关闭
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape') {
                modal.remove();
                overlay.remove();
            }
        });
    }
    
    // 从输入框开始搜索
    function startFromInput() {
        const input = document.getElementById('searchDataInput');
        const data = input.value.trim();
        
        if (!data) {
            alert('请输入新发展团员编号');
            input.focus();
            return;
        }
        
        // 解析数据
        const items = data.split('\n')
            .map(item => item.trim())
            .filter(item => item.length > 0);
        
        if (items.length === 0) {
            alert('没有有效的编号，请重新输入');
            input.focus();
            return;
        }
        
        // 移除模态框
        document.getElementById('dataInputModal').remove();
        document.getElementById('modalOverlay').remove();
        
        // 开始搜索
        startSearchProcess(items);
    }
    
    // 开始搜索流程
    function startSearchProcess(items) {
        if (isRunning) {
            alert('搜索正在进行中，请等待完成');
            return;
        }
        
        searchItems = items;
        
        console.log(`开始批量搜索，共 ${searchItems.length} 个新发展团员编号`);
        alert(`开始批量搜索 ${searchItems.length} 个新发展团员编号\n\n请保持页面打开，不要操作浏览器。`);
        
        isRunning = true;
        allResults = [];
        searchIndex = 0;
        
        // 添加控制面板
        addControlPanel(searchItems.length);
        
        // 开始搜索
        searchNextItem();
    }
    
    // 查找新发展团员编号输入框
    function findSearchInput() {
        // 查找包含"新发展团员编号"的输入框
        const inputs = document.querySelectorAll('input[placeholder*="新发展团员编号"]');
        if (inputs.length > 0) {
            return inputs[0];
        }
        
        // 备用方案：查找所有输入框
        const allInputs = document.querySelectorAll('input[type="text"]');
        for (let input of allInputs) {
            if (input.placeholder && input.placeholder.includes('新发展团员编号')) {
                return input;
            }
        }
        
        console.error('找不到新发展团员编号输入框');
        return null;
    }
    
    // 查找搜索按钮
    function findSearchButton() {
        // 方法1: 查找包含"搜索"文本的按钮
        const buttons = document.querySelectorAll('button');
        for (let button of buttons) {
            const span = button.querySelector('span');
            if (span && span.textContent.trim() === '搜索') {
                return button;
            }
        }
        
        // 方法2: 查找特定样式的按钮
        const searchButtons = document.querySelectorAll('button.el-button--primary');
        for (let button of searchButtons) {
            const span = button.querySelector('span');
            if (span && span.textContent.includes('搜索')) {
                return button;
            }
        }
        
        console.error('找不到搜索按钮');
        return null;
    }
    
    // 检查是否有错误弹窗
    function hasErrorPopup() {
        // 查找错误提示弹窗
        const errorPopup = document.querySelector('.el-message--error');
        if (errorPopup) {
            // 获取错误信息文本
            const errorContent = errorPopup.querySelector('.el-message__content');
            const errorText = errorContent ? errorContent.textContent.trim() : '搜索失败';
            
            // 关闭弹窗（点击关闭按钮）
            const closeBtn = errorPopup.querySelector('.el-message__closeBtn');
            if (closeBtn) {
                closeBtn.click();
            }
            
            return {
                hasError: true,
                errorText: errorText
            };
        }
        
        return { hasError: false, errorText: '' };
    }
    
    // 检查是否有结果表格
    function hasResultTable() {
        // 查找结果表格
        const resultTable = document.querySelector('.el-table__body tbody');
        if (resultTable && resultTable.children.length > 0) {
            return { hasTable: true, rows: resultTable.children.length };
        }
        
        // 检查是否有"暂无数据"的提示
        const emptyText = document.querySelector('.el-table__empty-text');
        if (emptyText) {
            return { hasTable: false, isEmpty: true };
        }
        
        return { hasTable: false, isEmpty: false };
    }
    
    // 解析搜索结果表格
    function parseSearchResults() {
        const results = [];
        
        // 查找结果表格
        const resultRows = document.querySelectorAll('.el-table__row');
        
        if (resultRows.length === 0) {
            return { hasResults: false, results: [] };
        }
        
        console.log(`找到 ${resultRows.length} 行结果`);
        
        // 解析每一行
        resultRows.forEach(row => {
            const result = {};
            const cells = row.querySelectorAll('td');
            
            // 根据搜索结果页面的表格结构解析
            if (cells.length >= 5) {
                // 最常见的5列表格结构
                result.姓名 = cells[0]?.textContent?.trim() || '';
                result.手机号码 = cells[1]?.textContent?.trim() || '';
                result.职务 = cells[2]?.textContent?.trim() || '';
                result.所在团组织 = cells[3]?.textContent?.trim() || '';
                result.操作 = cells[4]?.textContent?.trim() || '';
            } else if (cells.length >= 3) {
                // 简化的3列表格结构
                for (let i = 0; i < cells.length; i++) {
                    result[`列${i + 1}`] = cells[i]?.textContent?.trim() || '';
                }
            } else if (cells.length > 0) {
                // 其他表格结构
                cells.forEach((cell, index) => {
                    result[`列${index + 1}`] = cell.textContent.trim();
                });
            }
            
            if (Object.keys(result).length > 0) {
                results.push(result);
            }
        });
        
        return { hasResults: results.length > 0, results };
    }
    
    // 搜索下一个项目
    function searchNextItem() {
        if (!isRunning || searchIndex >= searchItems.length) {
            finishSearch();
            return;
        }
        
        const searchTerm = searchItems[searchIndex];
        updateStatus(`正在搜索编号: ${searchTerm} (${searchIndex + 1}/${searchItems.length})`);
        
        console.log(`搜索新发展团员编号: ${searchTerm}`);
        
        // 查找并填写搜索框
        const searchInput = findSearchInput();
        if (!searchInput) {
            console.error('找不到搜索输入框');
            finishSearch();
            return;
        }
        
        // 清空输入框
        searchInput.value = '';
        searchInput.dispatchEvent(new Event('input', { bubbles: true }));
        
        // 设置搜索词
        setTimeout(() => {
            searchInput.value = searchTerm;
            searchInput.dispatchEvent(new Event('input', { bubbles: true }));
            searchInput.dispatchEvent(new Event('change', { bubbles: true }));
            
            console.log(`已输入编号: ${searchTerm}`);
            
            // 点击搜索按钮
            setTimeout(() => {
                const searchButton = findSearchButton();
                
                if (searchButton && !searchButton.disabled) {
                    console.log('点击搜索按钮');
                    searchButton.click();
                    
                    // 等待结果或错误弹窗
                    setTimeout(() => {
                        checkSearchResult(searchTerm);
                    }, 1500);
                } else {
                    console.error('找不到搜索按钮或按钮被禁用');
                    recordErrorResult(searchTerm, '搜索按钮无效');
                }
            }, 800);
        }, 500);
    }
    
    // 检查搜索结果
    function checkSearchResult(searchTerm) {
        // 首先检查是否有错误弹窗
        const errorInfo = hasErrorPopup();
        if (errorInfo.hasError) {
            console.log(`搜索失败: ${searchTerm}, 错误信息: ${errorInfo.errorText}`);
            
            // 记录错误结果
            recordErrorResult(searchTerm, errorInfo.errorText);
            return;
        }
        
        // 检查是否有结果表格
        const tableInfo = hasResultTable();
        
        if (tableInfo.hasTable) {
            // 有结果表格，解析结果
            const parseResult = parseSearchResults();
            
            if (parseResult.hasResults && parseResult.results.length > 0) {
                // 有搜索结果
                parseResult.results.forEach(result => {
                    const fullResult = {
                        序号: searchIndex + 1,
                        新发展团员编号: searchTerm,
                        状态: '成功',
                        错误信息: '',
                        搜索时间: new Date().toLocaleString(),
                        ...result
                    };
                    allResults.push(fullResult);
                });
                console.log(`找到 ${parseResult.results.length} 条结果: ${searchTerm}`);
            } else {
                // 表格存在但没有结果
                recordErrorResult(searchTerm, '未找到（空表格）');
            }
        } else if (tableInfo.isEmpty) {
            // 有"暂无数据"提示
            recordErrorResult(searchTerm, '未找到');
        } else {
            // 既没有错误弹窗也没有表格，可能还在加载
            console.log('等待页面响应...');
            
            // 再次检查
            setTimeout(() => {
                checkSearchResult(searchTerm);
            }, 500);
            return;
        }
        
        // 完成当前搜索项
        finishCurrentItem();
    }
    
    // 记录错误结果
    function recordErrorResult(searchTerm, errorMessage) {
        const errorResult = {
            序号: searchIndex + 1,
            新发展团员编号: searchTerm,
            状态: '失败',
            错误信息: errorMessage,
            搜索时间: new Date().toLocaleString()
        };
        allResults.push(errorResult);
        
        // 完成当前搜索项
        finishCurrentItem();
    }
    
    // 完成当前搜索项
    function finishCurrentItem() {
        // 更新进度
        updateProgress(searchIndex + 1, searchItems.length);
        
        // 继续下一个搜索
        searchIndex++;
        
        // 短暂延迟后搜索下一个
        setTimeout(() => {
            searchNextItem();
        }, 1000);
    }
    
    // 完成搜索
    function finishSearch() {
        isRunning = false;
        updateStatus('搜索完成，正在生成结果文件...');
        
        // 按照搜索顺序排序结果
        sortResultsBySearchOrder();
        
        // 生成结果文件
        const timestamp = new Date().getTime();
        const dateStr = new Date().toISOString().split('T')[0].replace(/-/g, '');
        const filename = `新发展团员搜索结果_${dateStr}_${timestamp}`;
        
        if (config.saveAsCSV) {
            downloadCSV(filename);
        } else {
            downloadJSON(filename);
        }
        
        // 显示统计信息
        const successCount = allResults.filter(r => r.状态 === '成功').length;
        const failCount = allResults.filter(r => r.状态 === '失败').length;
        
        const summary = `
搜索完成！
总搜索数: ${searchItems.length}
成功找到: ${successCount} 条记录
搜索失败: ${failCount} 个编号
        `;
        
        // 在控制台显示结果摘要
        console.log('=== 搜索结果摘要 ===');
        console.log(`总搜索数: ${searchItems.length}`);
        console.log(`成功找到: ${successCount} 条记录`);
        console.log(`搜索失败: ${failCount} 个编号`);
        
        if (failCount > 0) {
            console.log('\n搜索失败的编号:');
            allResults.filter(r => r.状态 === '失败').forEach(item => {
                console.log(`  - ${item.新发展团员编号}: ${item.错误信息}`);
            });
        }
        
        // 显示结果文件信息
        setTimeout(() => {
            alert(`搜索完成！\n总搜索数: ${searchItems.length}\n成功找到: ${successCount} 条记录\n搜索失败: ${failCount} 个编号\n\n结果文件已开始下载。`);
        }, 1000);
        
        // 隐藏控制面板
        hideControlPanel();
    }
    
    // 按照搜索顺序排序结果
    function sortResultsBySearchOrder() {
        // 按照序号排序
        allResults.sort((a, b) => a.序号 - b.序号);
    }
    
    // 下载CSV文件
    function downloadCSV(filename) {
        let csvContent = '';
        
        if (allResults.length > 0) {
            // 获取所有字段名
            const allFields = new Set();
            allResults.forEach(result => {
                Object.keys(result).forEach(key => allFields.add(key));
            });
            
            // 定义固定的字段顺序
            const fieldOrder = [
                '序号',
                '新发展团员编号',
                '状态',
                '错误信息',
                '姓名',
                '手机号码',
                '职务',
                '所在团组织',
                '搜索时间'
            ];
            
            // 优先使用固定顺序的字段，然后添加其他字段
            const headers = [];
            fieldOrder.forEach(field => {
                if (allFields.has(field)) {
                    headers.push(field);
                    allFields.delete(field);
                }
            });
            
            // 添加剩余字段
            Array.from(allFields).forEach(field => {
                headers.push(field);
            });
            
            // 添加表头
            csvContent += headers.join(',') + '\n';
            
            // 添加数据
            allResults.forEach(result => {
                const row = headers.map(header => {
                    let value = result[header] || '';
                    // 处理包含逗号、引号的内容
                    if (typeof value === 'string' && (value.includes(',') || value.includes('"'))) {
                        value = `"${value.replace(/"/g, '""')}"`;
                    }
                    return value;
                });
                csvContent += row.join(',') + '\n';
            });
            
            // 添加搜索统计
            const successCount = allResults.filter(r => r.状态 === '成功').length;
            const failCount = allResults.filter(r => r.状态 === '失败').length;
            
            csvContent += `\n\n=== 搜索统计 ===\n项目,数量\n总搜索数,${searchItems.length}\n成功找到,${successCount}\n搜索失败,${failCount}\n开始时间,${allResults.length > 0 ? allResults[0].搜索时间 : '无'}\n完成时间,${new Date().toLocaleString()}`;
        } else {
            csvContent = '序号,新发展团员编号,状态,错误信息,搜索时间\n(无搜索结果)';
        }
        
        // 创建下载链接
        const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        
        link.setAttribute('href', url);
        link.setAttribute('download', `${filename}.csv`);
        link.style.visibility = 'hidden';
        
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        console.log('CSV文件已生成并开始下载');
    }
    
    // 添加控制面板
    function addControlPanel(total) {
        let panel = document.getElementById('searchControlPanel');
        
        if (!panel) {
            panel = document.createElement('div');
            panel.id = 'searchControlPanel';
            panel.style.cssText = `
                position: fixed;
                top: 20px;
                right: 20px;
                background: white;
                border: 2px solid #409EFF;
                border-radius: 8px;
                padding: 15px;
                z-index: 999999;
                box-shadow: 0 2px 12px rgba(0,0,0,0.1);
                min-width: 300px;
                font-family: Arial, sans-serif;
            `;
            
            panel.innerHTML = `
                <div style="margin-bottom: 10px; font-weight: bold; color: #409EFF;">
                    新发展团员编号批量搜索
                </div>
                <div id="searchStatus" style="margin-bottom: 10px; color: #666;">
                    准备开始...
                </div>
                <div style="margin-bottom: 10px;">
                    <div style="background: #f5f5f5; border-radius: 4px; overflow: hidden;">
                        <div id="searchProgressBar" style="height: 20px; background: #409EFF; width: 0%; transition: width 0.3s;"></div>
                    </div>
                    <div id="searchProgressText" style="text-align: center; margin-top: 5px; font-size: 12px;">
                        0/${total}
                    </div>
                </div>
                <div style="display: flex; gap: 10px; margin-top: 15px;">
                    <button id="pauseSearch" style="flex: 1; padding: 8px; background: #f56c6c; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;">
                        暂停
                    </button>
                    <button id="stopSearch" style="flex: 1; padding: 8px; background: #909399; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;">
                        停止
                    </button>
                </div>
            `;
            
            document.body.appendChild(panel);
            
            // 添加事件监听
            document.getElementById('pauseSearch').addEventListener('click', togglePause);
            document.getElementById('stopSearch').addEventListener('click', stopSearch);
        }
    }
    
    // 更新状态
    function updateStatus(status) {
        const statusEl = document.getElementById('searchStatus');
        if (statusEl) {
            statusEl.textContent = status;
        }
    }
    
    // 更新进度
    function updateProgress(current, total) {
        const percentage = Math.round((current / total) * 100);
        const progressBar = document.getElementById('searchProgressBar');
        const progressText = document.getElementById('searchProgressText');
        
        if (progressBar) {
            progressBar.style.width = `${percentage}%`;
        }
        if (progressText) {
            progressText.textContent = `${current}/${total} (${percentage}%)`;
        }
    }
    
    // 切换暂停状态
    function togglePause() {
        const btn = document.getElementById('pauseSearch');
        if (isRunning) {
            isRunning = false;
            btn.textContent = '继续';
            btn.style.background = '#67c23a';
            updateStatus('已暂停 - 点击"继续"按钮恢复');
        } else {
            isRunning = true;
            btn.textContent = '暂停';
            btn.style.background = '#f56c6c';
            updateStatus('已恢复搜索');
            searchNextItem();
        }
    }
    
    // 停止搜索
    function stopSearch() {
        isRunning = false;
        updateStatus('已停止');
        hideControlPanel();
        
        // 如果已经收集了一些结果，询问是否保存
        if (allResults.length > 0) {
            if (confirm(`搜索已停止。\n已收集 ${allResults.length} 条记录。\n是否保存已收集的结果？`)) {
                // 按照搜索顺序排序结果
                sortResultsBySearchOrder();
                
                const timestamp = new Date().getTime();
                const filename = `新发展团员搜索结果_部分_${timestamp}`;
                downloadCSV(filename);
            }
        } else {
            alert('搜索已停止，没有收集到结果。');
        }
    }
    
    // 隐藏控制面板
    function hideControlPanel() {
        const panel = document.getElementById('searchControlPanel');
        if (panel) {
            panel.style.display = 'none';
        }
    }
    
    // 显示使用说明
    function showInstructions() {
        const instructions = `
=== 新发展团员编号批量搜索工具（顺序版）===

特点：
1. 结果按照您输入的顺序排列
2. 搜索成功和失败的结果都在同一个表格中
3. 搜索失败的在表格中用错误信息标记
4. 不再单独分开展示未找到的记录

使用方法：
1. 复制Excel文件中的新发展团员编号（每行一个）
2. 在新发展团员搜索页面（页面必须完全加载）
3. 按F12打开开发者工具
4. 进入Console（控制台）标签页
5. 粘贴此代码并按回车运行
6. 在弹出的窗口中粘贴编号并开始搜索

注意事项：
- 请勿在脚本运行期间操作浏览器
- 确保网络连接稳定
- 搜索结果会自动保存为CSV文件
- 搜索间隔设置为2秒，避免请求过快
- 如果遇到问题，可以在控制台中输入 window.newMemberSearch.stop() 停止

点击确定继续...
        `;
        
        if (confirm(instructions)) {
            createDataInputModal();
        }
    }
    
    // 初始化
    console.clear();
    console.log('新发展团员编号批量搜索工具（GitHub CDN版）已加载');
    console.log('特点：结果按照输入顺序排列，成功失败都在同一表格中');
    
    // 检查是否在正确的页面
    const searchInput = findSearchInput();
    if (!searchInput) {
        console.warn('警告：未找到新发展团员编号输入框，请确保在当前页面运行此脚本');
    }
    
    showInstructions();
    
    // 暴露一些函数到全局，便于调试
    window.newMemberSearch = {
        start: () => createDataInputModal(),
        pause: togglePause,
        stop: stopSearch,
        getResults: () => allResults,
        config
    };
    
})();
