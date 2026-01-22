// 全局变量
let uploadedFile = null;
let parsedData = [];
let chart = null;
let currentHoleNumber = null;

// DOM元素
const fileInput = document.getElementById('fileInput');
const previewSection = document.getElementById('previewSection');
const dataTable = document.getElementById('dataTable');
const calculateSection = document.getElementById('calculateSection');
const calculateResult = document.getElementById('calculateResult');
const chartSection = document.getElementById('chartSection');
const chartCanvas = document.getElementById('chartCanvas');
const loadingOverlay = document.getElementById('loadingOverlay');
const toast = document.getElementById('toast');
const toastMessage = document.getElementById('toastMessage');

// 常量定义
const COLUMN_INDEXES = {
    HOLE_NUMBER: 2,     // 孔号在第3列
    X_AXIS: 4,          // X轴为第5列锤击数
    DEPTH: 5,           // Y轴为第6列深度
    TOTAL_DEPTH: 6      // Y轴为第7列总深度
};

// 显示加载状态
function showLoading() {
    if (loadingOverlay) {
        loadingOverlay.style.display = 'flex';
    }
}

// 公共函数

/**
 * 获取指定列索引对应的列名
 * @param {Array} headers - 表头数组
 * @param {number} index - 列索引
 * @returns {string} 列名
 */
function getColumnName(headers, index) {
    return headers.length > index ? headers[index] : '';
}

/**
 * 过滤指定孔号的数据
 * @param {Array} data - 原始数据
 * @param {string} holeNumber - 孔号
 * @param {string} holeNumberColumn - 孔号列名
 * @returns {Array} 过滤后的数据
 */
function filterDataByHoleNumber(data, holeNumber, holeNumberColumn) {
    if (holeNumber === 'all') {
        return data;
    } else {
        return data.filter(row => row[holeNumberColumn] === holeNumber);
    }
}

/**
 * 生成唯一的孔号列表
 * @param {Array} data - 原始数据
 * @param {string} holeNumberColumn - 孔号列名
 * @returns {Array} 唯一的孔号列表
 */
function generateUniqueHoleNumbers(data, holeNumberColumn) {
    return [...new Set(data.map(row => row[holeNumberColumn]))].filter(holeNumber => 
        holeNumber !== undefined && holeNumber !== null && holeNumber !== ''
    );
}

// 隐藏加载状态
function hideLoading() {
    if (loadingOverlay) {
        loadingOverlay.style.display = 'none';
    }
}

// 显示Toast通知
/**
 * 显示Toast通知
 * @param {string} message - 通知消息
 * @param {string} type - 通知类型：'success' | 'error' | 'warning' | 'info'
 * @param {number} duration - 显示持续时间（毫秒），默认3000
 */
function showToast(message, type = 'info', duration = 3000) {
    if (!toast || !toastMessage) return;
    
    // 设置消息内容
    toastMessage.textContent = message;
    
    // 移除之前的类型类
    toast.classList.remove('success', 'error', 'warning', 'info');
    
    // 添加新的类型类
    toast.classList.add(type);
    
    // 显示Toast
    toast.style.display = 'block';
    setTimeout(() => {
        toast.classList.add('show');
    }, 10);
    
    // 自动隐藏
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => {
            toast.style.display = 'none';
        }, 300);
    }, duration);
}

// 初始化应用
function initApp() {
    initEventListeners();
    console.log('应用初始化完成');
}

// 初始化事件监听
function initEventListeners() {
    // 文件选择事件
    fileInput.addEventListener('change', handleFileSelect);
    
    // 导入数据按钮事件
    document.getElementById('importButton').addEventListener('click', () => {
        fileInput.click();
    });
    
    // 清空重置按钮事件
    document.getElementById('resetButton').addEventListener('click', resetApp);
    

    

    
    // 导出按钮事件
    document.getElementById('exportImageButton').addEventListener('click', exportChartAsImage);
    document.getElementById('exportExcelButton').addEventListener('click', exportResultAsExcel);
}

// 文件处理相关函数

// 处理文件选择
function handleFileSelect(e) {
    if (e.target.files.length) {
        showLoading();
        uploadedFile = e.target.files[0];
        processFile(uploadedFile);
    }
}

// 处理文件
function processFile(file) {
    const fileName = file.name;
    const fileExtension = fileName.split('.').pop().toLowerCase();
    
    if (fileExtension === 'csv') {
        parseCSV(file);
    } else if (fileExtension === 'xlsx') {
        parseXLSX(file);
    } else {
        showToast('文件格式错误\n\n请上传以下格式的文件：\n- CSV文件 (.csv)\n- Excel文件 (.xlsx)\n\n当前文件格式：.' + fileExtension, 'error');
        hideLoading();
    }
}

// 解析CSV文件
function parseCSV(file) {
    // 优先尝试GBK编码，因为中文Windows系统默认使用GBK
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const text = e.target.result;
            console.log('使用GBK编码读取文件成功');
            console.log('文件前100个字符:', text.substring(0, 100));
            
            // 使用Papa Parse解析CSV文本
            Papa.parse(text, {
                header: true,
                complete: function(results) {
                    console.log('Papa Parse解析结果:', results.data);
                    parsedData = results.data;
                    displayData(parsedData);
                    showSections();
                    hideLoading();
                    showToast('CSV文件导入成功！', 'success');
                },
                error: function(error) {
                    console.error('CSV解析错误:', error);
                    showToast('CSV文件解析失败\n\n可能的原因：\n- 文件格式不正确\n- 文件编码错误\n- 文件内容为空或损坏\n\n请检查文件后重试', 'error');
                    hideLoading();
                }
            });
        } catch (error) {
            console.error('GBK编码读取错误:', error);
            // 如果GBK失败，尝试UTF-8
            const utf8Reader = new FileReader();
            utf8Reader.onload = function(e) {
                try {
                    const text = e.target.result;
                    console.log('使用UTF-8编码读取文件成功');
                    console.log('文件前100个字符:', text.substring(0, 100));
                    
                    // 使用Papa Parse解析CSV文本
                    Papa.parse(text, {
                        header: true,
                        complete: function(results) {
                            console.log('Papa Parse解析结果:', results.data);
                            parsedData = results.data;
                            displayData(parsedData);
                            showSections();
                            hideLoading();
                            showToast('CSV文件导入成功！', 'success');
                        },
                        error: function(error) {
                            console.error('CSV解析错误:', error);
                            showToast('CSV文件解析失败\n\n可能的原因：\n- 文件格式不正确\n- 文件编码错误\n- 文件内容为空或损坏\n\n请检查文件后重试', 'error');
                            hideLoading();
                        }
                    });
                } catch (error) {
                    console.error('UTF-8编码读取错误:', error);
                    showToast('文件读取失败\n\n可能的原因：\n- 文件编码错误\n- 文件格式不正确\n- 文件内容为空或损坏\n\n请检查文件后重试', 'error');
                    hideLoading();
                }
            };
            utf8Reader.onerror = function() {
                console.error('UTF-8文件读取错误');
                showToast('文件读取失败\n\n可能的原因：\n- 文件编码错误\n- 文件格式不正确\n- 文件内容为空或损坏\n\n请检查文件后重试', 'error');
                hideLoading();
            };
            utf8Reader.readAsText(file, 'utf-8');
        }
    };
    reader.onerror = function() {
        console.error('GBK文件读取错误');
        showToast('文件读取失败\n\n可能的原因：\n- 文件编码错误\n- 文件格式不正确\n- 文件内容为空或损坏\n\n请检查文件后重试', 'error');
        hideLoading();
    };
    reader.readAsText(file, 'gbk');
}

// 解析XLSX文件
function parseXLSX(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        parsedData = XLSX.utils.sheet_to_json(worksheet);
        displayData(parsedData);
        showSections();
        hideLoading();
        showToast('Excel文件导入成功！', 'success');
    };
    reader.onerror = function() {
        console.error('XLSX文件读取错误');
        showToast('Excel文件读取失败\n\n可能的原因：\n- 文件格式不正确\n- 文件版本过旧（请使用.xlsx格式）\n- 文件内容为空或损坏\n- 文件受到保护或加密\n\n请检查文件后重试', 'error');
        hideLoading();
    };
    reader.readAsArrayBuffer(file);
}

// 界面显示相关函数

// 为表格添加数据条
function addDataBars() {
    // 获取表格元素
    const table = document.getElementById('dataTable');
    if (!table) return;
    
    const tbody = table.querySelector('tbody');
    if (!tbody) return;
    
    // 获取所有行
    const rows = Array.from(tbody.querySelectorAll('tr'));
    if (rows.length === 0) return;
    
    // 确定深度和总深度列的索引
    const headers = table.querySelectorAll('thead th');
    let depthColumnIndex = -1;
    let totalDepthColumnIndex = -1;
    
    // 缓存列索引
    for (let i = 0; i < headers.length; i++) {
        const text = headers[i].textContent.trim();
        if (text === '深度') {
            depthColumnIndex = i;
        } else if (text === '总深度') {
            totalDepthColumnIndex = i;
        }
    }
    
    if (depthColumnIndex === -1 || totalDepthColumnIndex === -1) return;
    
    // 计算深度和总深度的最大值
    let maxDepth = 0;
    let maxTotalDepth = 0;
    
    // 第一次遍历：计算最大值
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const cells = row.querySelectorAll('td');
        
        if (cells.length > depthColumnIndex) {
            const depthValue = parseFloat(cells[depthColumnIndex].textContent);
            if (!isNaN(depthValue) && depthValue > maxDepth) {
                maxDepth = depthValue;
            }
        }
        
        if (cells.length > totalDepthColumnIndex) {
            const totalDepthValue = parseFloat(cells[totalDepthColumnIndex].textContent);
            if (!isNaN(totalDepthValue) && totalDepthValue > maxTotalDepth) {
                maxTotalDepth = totalDepthValue;
            }
        }
    }
    
    // 第二次遍历：添加数据条
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const cells = row.querySelectorAll('td');
        
        // 处理深度列
        if (cells.length > depthColumnIndex) {
            const depthCell = cells[depthColumnIndex];
            const depthValue = parseFloat(depthCell.textContent);
            
            if (!isNaN(depthValue) && maxDepth > 0) {
                const percentage = (depthValue / maxDepth) * 100;
                depthCell.style.position = 'relative';
                depthCell.style.paddingRight = '80px';
                
                // 根据是否是最大值选择不同的渐变颜色
                let gradientColor;
                if (depthValue === maxDepth) {
                    // 最大值使用橙色渐变
                    gradientColor = 'linear-gradient(90deg, rgba(255, 165, 0, 0.2), rgba(255, 165, 0, 0.6))';
                } else {
                    // 其他值使用蓝色渐变
                    gradientColor = 'linear-gradient(90deg, rgba(52, 152, 219, 0.2), rgba(52, 152, 219, 0.6))';
                }
                
                // 构建HTML内容
                depthCell.innerHTML = `
                    <span style="position: relative; z-index: 2;">${depthValue}</span>
                    <span style="
                        position: absolute;
                        left: 0;
                        top: 0;
                        height: 100%;
                        width: ${percentage}%;
                        background: ${gradientColor};
                        z-index: 1;
                        border-radius: 3px;
                    "></span>
                `;
            }
        }
        
        // 处理总深度列
        if (cells.length > totalDepthColumnIndex) {
            const totalDepthCell = cells[totalDepthColumnIndex];
            const totalDepthValue = parseFloat(totalDepthCell.textContent);
            
            if (!isNaN(totalDepthValue) && maxTotalDepth > 0) {
                const percentage = (totalDepthValue / maxTotalDepth) * 100;
                totalDepthCell.style.position = 'relative';
                totalDepthCell.style.paddingRight = '80px';
                
                // 构建HTML内容
                totalDepthCell.innerHTML = `
                    <span style="position: relative; z-index: 2;">${totalDepthValue}</span>
                    <span style="
                        position: absolute;
                        left: 0;
                        top: 0;
                        height: 100%;
                        width: ${percentage}%;
                        background: linear-gradient(90deg, rgba(46, 204, 113, 0.2), rgba(46, 204, 113, 0.6));
                        z-index: 1;
                        border-radius: 3px;
                    "></span>
                `;
            }
        }
    }
}

// 显示数据
function displayData(data) {
    if (!data || data.length === 0) {
        dataTable.innerHTML = '<tr><td colspan="100%">没有数据</td></tr>';
        return;
    }
    
    const headers = Object.keys(data[0]);
    let html = '<thead><tr>';
    headers.forEach(header => {
        html += `<th>${header}</th>`;
    });
    html += '</tr></thead><tbody>';
    
    // 显示所有数据
    data.forEach(row => {
        html += '<tr>';
        headers.forEach(header => {
            html += `<td>${row[header] || ''}</td>`;
        });
        html += '</tr>';
    });
    
    html += '</tbody>';
    dataTable.innerHTML = html;
    
    // 添加数据条
    addDataBars();
}



// 显示后续区域
function showSections() {
    previewSection.style.display = 'block';
    calculateSection.style.display = 'block';
    chartSection.style.display = 'block';
    
    // 初始化图表
    initChart();
    
    // 生成孔号按钮
    generateHoleButtons();
    
    // 自动生成深度统计结果
    generateDepthStatistics();
}

// 生成孔号按钮
function generateHoleButtons() {
    const holeSelector = document.getElementById('holeSelector');
    const holeButtonsContainer = document.getElementById('holeButtons');
    
    // 清空现有按钮
    holeButtonsContainer.innerHTML = '';
    
    if (parsedData.length > 0) {
        const headers = Object.keys(parsedData[0]);
        
        // 检查列索引是否有效
        if (headers.length > COLUMN_INDEXES.HOLE_NUMBER) {
            const holeNumberColumn = getColumnName(headers, COLUMN_INDEXES.HOLE_NUMBER);
            
            // 获取唯一的孔号列表
            const uniqueHoles = generateUniqueHoleNumbers(parsedData, holeNumberColumn);

            
            if (uniqueHoles.length > 0) {
                // 显示孔号选择区域
                holeSelector.style.display = 'block';
                
                // 创建"全部"按钮
                const allButton = document.createElement('button');
                allButton.textContent = '全部';
                allButton.dataset.holeNumber = 'all';
                
                // 添加点击事件
                allButton.addEventListener('click', () => {
                    // 切换活动状态
                    document.querySelectorAll('.hole-buttons button').forEach(btn => {
                        btn.classList.remove('active');
                    });
                    allButton.classList.add('active');
                    
                    // 显示所有孔号的数据
                    showHoleData('all');
                });
                
                holeButtonsContainer.appendChild(allButton);
                
                // 为每个孔号创建按钮
                uniqueHoles.forEach(holeNumber => {
                    const button = document.createElement('button');
                    button.textContent = `孔${holeNumber}`;
                    button.dataset.holeNumber = holeNumber;
                    
                    // 添加点击事件
                    button.addEventListener('click', () => {
                        // 切换活动状态
                        document.querySelectorAll('.hole-buttons button').forEach(btn => {
                            btn.classList.remove('active');
                        });
                        button.classList.add('active');
                        
                        // 显示对应孔号的数据
                        showHoleData(holeNumber);
                    });
                    
                    holeButtonsContainer.appendChild(button);
                });
            }
        }
    }
}

// 图表相关函数

/**
 * 初始化图表
 * @description 创建并初始化Chart.js图表实例
 */
function initChart() {
    const ctx = chartCanvas.getContext('2d');
    
    // 准备图表数据
    const datasets = [];
    
    // 按孔号分组数据
    if (parsedData.length > 0) {
        const headers = Object.keys(parsedData[0]);
        
        // 检查列索引是否有效
        if (headers.length > COLUMN_INDEXES.TOTAL_DEPTH) {
            // 获取各列名
            const holeNumberColumn = getColumnName(headers, COLUMN_INDEXES.HOLE_NUMBER);
            const xAxisColumn = getColumnName(headers, COLUMN_INDEXES.X_AXIS);
            const depthColumn = getColumnName(headers, COLUMN_INDEXES.DEPTH);
            const totalDepthColumn = getColumnName(headers, COLUMN_INDEXES.TOTAL_DEPTH);
            
            console.log('列信息:', {
                holeNumberColumn,
                xAxisColumn,
                depthColumn,
                totalDepthColumn
            });
            
            // 使用X轴列作为标签
            const labels = parsedData.map(row => {
                const value = row[xAxisColumn];
                return value !== undefined ? value : '';
            });
            
            // 获取第一个孔号
            const firstHoleNumber = parsedData[0][holeNumberColumn];
            console.log('第一个孔号:', firstHoleNumber);
            
            // 只为第一个孔号创建数据集
            const colors = ['rgba(52, 152, 219, 0.8)', 'rgba(46, 204, 113, 0.8)'];
            
            // 为深度创建数据集
            const depthData = parsedData.map(row => {
                if (row[holeNumberColumn] === firstHoleNumber) {
                    const value = parseFloat(row[depthColumn]);
                    return isNaN(value) ? null : value;
                } else {
                    return null;
                }
            });
            
            // 为总深度创建数据集
            const totalDepthData = parsedData.map(row => {
                if (row[holeNumberColumn] === firstHoleNumber) {
                    const value = parseFloat(row[totalDepthColumn]);
                    return isNaN(value) ? null : value;
                } else {
                    return null;
                }
            });
            
            // 添加深度数据集
            datasets.push({
                label: `${firstHoleNumber}-${depthColumn}`,
                data: depthData,
                borderColor: colors[0],
                backgroundColor: colors[0].replace('0.8', '0.2'),
                tension: 0.1
            });
            
            // 添加总深度数据集
            datasets.push({
                label: `${firstHoleNumber}-${totalDepthColumn}`,
                data: totalDepthData,
                borderColor: colors[1],
                backgroundColor: colors[1].replace('0.8', '0.2'),
                tension: 0.1
            });
            
            console.log('生成的数据集:', datasets);
            
            // 创建图表（默认不显示曲线）
            chart = new Chart(ctx, {
                type: 'line',
                data: {
                    labels: labels,
                    datasets: []
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: true,
                            text: `孔${firstHoleNumber}数据折线图表`
                        }
                    },
                    scales: {
                        x: {
                            title: {
                                display: true,
                                text: '锤击数'
                            }
                        },
                        y: {
                            type: 'linear',
                            display: true,
                            position: 'left',
                            title: {
                                display: true,
                                text: '深度'
                            },
                            beginAtZero: true
                        },
                        y1: {
                            type: 'linear',
                            display: true,
                            position: 'right',
                            title: {
                                display: true,
                                text: '总深度'
                            },
                            beginAtZero: true,
                            grid: {
                                drawOnChartArea: false
                            }
                        }
                    }
                }
            });
        } else {
            // 如果列索引无效，使用默认方法
            const labels = parsedData.map((row, index) => index + 1);
            
            headers.forEach((header, index) => {
                const data = parsedData.map(row => {
                    const value = parseFloat(row[header]);
                    return isNaN(value) ? null : value;
                });
                
                datasets.push({
                    label: header,
                    data: data,
                    borderColor: `rgba(${Math.random() * 255}, ${Math.random() * 255}, ${Math.random() * 255}, 0.8)`,
                    backgroundColor: `rgba(${Math.random() * 255}, ${Math.random() * 255}, ${Math.random() * 255}, 0.2)`,
                    tension: 0.1
                });
            });
            
            // 创建图表（默认不显示曲线）
            chart = new Chart(ctx, {
                type: 'line',
                data: {
                    labels: labels,
                    datasets: []
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: true,
                            text: '数据折线图表'
                        }
                    },
                    scales: {
                        x: {
                            title: {
                                display: true,
                                text: '数据点'
                            }
                        },
                        y: {
                            beginAtZero: true
                        }
                    }
                }
            });
        }
    }
}



// 显示对应孔号的数据
function showHoleData(holeNumber) {
    if (!parsedData || parsedData.length === 0) return;
    
    // 更新当前孔号
    currentHoleNumber = holeNumber;
    
    const headers = Object.keys(parsedData[0]);
    
    // 获取孔号列名
    const holeNumberColumn = getColumnName(headers, COLUMN_INDEXES.HOLE_NUMBER);
    
    // 筛选数据
    const holeData = filterDataByHoleNumber(parsedData, holeNumber, holeNumberColumn);
    
    // 更新数据预览
    updateDataPreview(holeData);
    
    // 更新深度统计结果
    generateDepthStatistics(holeNumber);
    
    // 只有当不是"全部"孔号时才更新图表
    if (holeNumber !== 'all') {
        // 更新图表
        updateChartForHole(holeNumber, holeData);
    }
}

// 更新数据预览
function updateDataPreview(data) {
    if (!data || data.length === 0) {
        dataTable.innerHTML = '<tr><td colspan="100%">没有数据</td></tr>';
        return;
    }
    
    const headers = Object.keys(data[0]);
    let html = '<thead><tr>';
    headers.forEach(header => {
        html += `<th>${header}</th>`;
    });
    html += '</tr></thead><tbody>';
    
    // 显示所有数据
    data.forEach(row => {
        html += '<tr>';
        headers.forEach(header => {
            html += `<td>${row[header] || ''}</td>`;
        });
        html += '</tr>';
    });
    
    html += '</tbody>';
    dataTable.innerHTML = html;
    
    // 添加数据条
    addDataBars();
}

// 更新计算结果
function updateCalculateResult(data) {
    if (!data || data.length === 0) {
        calculateResult.innerHTML = '没有数据可计算';
        return;
    }
    
    // 保留当前的计算结果，或者清空显示
    // 这里我们选择清空显示，让用户可以重新计算
    calculateResult.innerHTML = '请选择计算类型';
}

// 更新图表
function updateChartForHole(holeNumber, data) {
    if (!chart || !data || data.length === 0) return;
    
    const headers = Object.keys(data[0]);
    
    // 列索引定义（从0开始）
    const X_AXIS_INDEX = 4; // X轴为第5列锤击数
    const DEPTH_INDEX = 5; // Y轴为第6列深度
    const TOTAL_DEPTH_INDEX = 6; // Y轴为第7列宗深度
    const HOLE_NUMBER_INDEX = 2; // 孔号在第3列
    
    // 检查列索引是否有效
    if (headers.length > TOTAL_DEPTH_INDEX) {
        const xAxisColumn = headers[X_AXIS_INDEX];
        const depthColumn = headers[DEPTH_INDEX];
        const totalDepthColumn = headers[TOTAL_DEPTH_INDEX];
        const holeNumberColumn = headers[HOLE_NUMBER_INDEX];
        
        if (holeNumber === 'all') {
            // 处理"全部"孔号的情况
            // 按孔号分组数据
            const holeGroups = {};
            data.forEach(row => {
                const currentHoleNumber = row[holeNumberColumn];
                if (!holeGroups[currentHoleNumber]) {
                    holeGroups[currentHoleNumber] = [];
                }
                holeGroups[currentHoleNumber].push(row);
            });
            
            // 生成不同颜色
            const colors = [
                'rgba(52, 152, 219, 0.8)',  // 蓝色
                'rgba(46, 204, 113, 0.8)',  // 绿色
                'rgba(231, 76, 60, 0.8)',   // 红色
                'rgba(155, 89, 182, 0.8)',  // 紫色
                'rgba(52, 73, 94, 0.8)',    // 深灰色
                'rgba(241, 196, 15, 0.8)',  // 黄色
                'rgba(230, 126, 34, 0.8)',  // 橙色
                'rgba(149, 165, 166, 0.8)'  // 浅灰色
            ];
            
            // 为每个孔号创建数据集
            const datasets = [];
            let colorIndex = 0;
            
            Object.entries(holeGroups).forEach(([currentHoleNumber, holeData]) => {
                const color = colors[colorIndex % colors.length];
                const backgroundColor = color.replace('0.8', '0.2');
                
                // 为深度创建数据集
                const depthData = holeData.map(row => {
                    const value = parseFloat(row[depthColumn]);
                    return isNaN(value) ? null : value;
                });
                
                // 为总深度创建数据集
                const totalDepthData = holeData.map(row => {
                    const value = parseFloat(row[totalDepthColumn]);
                    return isNaN(value) ? null : value;
                });
                
                // 添加深度数据集
                datasets.push({
                    label: `孔${currentHoleNumber}-${depthColumn}`,
                    data: depthData,
                    borderColor: color,
                    backgroundColor: backgroundColor,
                    tension: 0.1,
                    yAxisID: 'y'
                });
                
                // 添加总深度数据集
                datasets.push({
                    label: `孔${currentHoleNumber}-${totalDepthColumn}`,
                    data: totalDepthData,
                    borderColor: color.replace('0.8', '1'),
                    backgroundColor: backgroundColor.replace('0.2', '0.1'),
                    tension: 0.1,
                    borderDash: [5, 5],
                    yAxisID: 'y1'
                });
                
                colorIndex++;
            });
            
            // 使用所有数据的X轴标签
            const labels = data.map(row => {
                const value = row[xAxisColumn];
                return value !== undefined ? value : '';
            });
            
            // 更新图表数据
            chart.data.labels = labels;
            chart.data.datasets = datasets;
            
            // 更新图表标题
            chart.options.plugins.title.text = '所有孔数据折线图表';
        } else {
            // 处理单个孔号的情况
            // 使用第5列作为X轴标签
            const labels = data.map(row => {
                const value = row[xAxisColumn];
                return value !== undefined ? value : '';
            });
            
            // 为深度创建数据集
            const depthData = data.map(row => {
                const value = parseFloat(row[depthColumn]);
                return isNaN(value) ? null : value;
            });
            
            // 为宗深度创建数据集
            const totalDepthData = data.map(row => {
                const value = parseFloat(row[totalDepthColumn]);
                return isNaN(value) ? null : value;
            });
            
            // 更新图表数据
            chart.data.labels = labels;
            chart.data.datasets = [
                {
                    label: `${holeNumber}-${depthColumn}`,
                    data: depthData,
                    borderColor: 'rgba(52, 152, 219, 0.8)',
                    backgroundColor: 'rgba(52, 152, 219, 0.2)',
                    tension: 0.1,
                    yAxisID: 'y'
                },
                {
                    label: `${holeNumber}-${totalDepthColumn}`,
                    data: totalDepthData,
                    borderColor: 'rgba(46, 204, 113, 0.8)',
                    backgroundColor: 'rgba(46, 204, 113, 0.2)',
                    tension: 0.1,
                    yAxisID: 'y1'
                }
            ];
            
            // 更新图表标题
            chart.options.plugins.title.text = `孔${holeNumber}数据折线图表`;
        }
        
        // 更新图表配置，确保双Y轴设置
        chart.options.scales = {
            x: {
                title: {
                    display: true,
                    text: '锤击数'
                }
            },
            y: {
                type: 'linear',
                display: true,
                position: 'left',
                title: {
                    display: true,
                    text: '深度'
                },
                beginAtZero: true
            },
            y1: {
                type: 'linear',
                display: true,
                position: 'right',
                title: {
                    display: true,
                    text: '总深度'
                },
                beginAtZero: true,
                grid: {
                    drawOnChartArea: false
                }
            }
        };
        
        // 更新图表
        chart.update();
    }
}

// 计算函数
function calculate(type) {
    if (!parsedData || parsedData.length === 0) {
        calculateResult.innerHTML = '没有数据可计算';
        return;
    }
    
    // 如果是深度统计，使用新的格式
    if (type === 'depth') {
        generateDepthStatistics();
        return;
    }
    
    // 获取要计算的数据
    const dataToCalculate = getCurrentDataForCalculate();
    
    if (!dataToCalculate || dataToCalculate.length === 0) {
        calculateResult.innerHTML = '没有数据可计算';
        return;
    }
    
    const headers = Object.keys(dataToCalculate[0]);
    let resultHTML = '<h3>计算结果:</h3><table>';
    
    headers.forEach(header => {
        const values = dataToCalculate.map(row => {
            const value = parseFloat(row[header]);
            return isNaN(value) ? null : value;
        }).filter(val => val !== null);
        
        if (values.length === 0) {
            resultHTML += `<tr><td>${header}</td><td>无有效数据</td></tr>`;
            return;
        }
        
        let result;
        switch (type) {
            case 'sum':
                result = values.reduce((acc, val) => acc + val, 0);
                break;
            case 'average':
                result = values.reduce((acc, val) => acc + val, 0) / values.length;
                break;
            case 'max':
                result = Math.max(...values);
                break;
            case 'min':
                result = Math.min(...values);
                break;
        }
        
        resultHTML += `<tr><td>${header}</td><td>${result.toFixed(2)}</td></tr>`;
    });
    
    resultHTML += '</table>';
    calculateResult.innerHTML = resultHTML;
}

// 生成深度统计结果
function generateDepthStatistics(holeNumber = null) {
    if (!parsedData || parsedData.length === 0) {
        calculateResult.innerHTML = '没有数据可计算';
        return;
    }
    
    const headers = Object.keys(parsedData[0]);
    const HOLE_NUMBER_INDEX = 2; // 孔号在第3列
    const HAMMER_COUNT_INDEX = 4; // 锤击数在第5列
    const DEPTH_INDEX = 5; // 深度在第6列
    const TOTAL_DEPTH_INDEX = 6; // 总深度在第7列
    
    const holeNumberColumn = headers[HOLE_NUMBER_INDEX];
    const hammerCountColumn = headers[HAMMER_COUNT_INDEX];
    const depthColumn = headers[DEPTH_INDEX];
    const totalDepthColumn = headers[TOTAL_DEPTH_INDEX];
    
    // 按孔号分组数据
    const holeGroups = {};
    parsedData.forEach(row => {
        const holeNumber = row[holeNumberColumn];
        // 过滤掉无效的孔号值
        if (holeNumber !== undefined && holeNumber !== null && holeNumber !== '') {
            if (!holeGroups[holeNumber]) {
                holeGroups[holeNumber] = [];
            }
            holeGroups[holeNumber].push(row);
        }
    });
    
    let resultHTML = '<div class="depth-statistics">';
    
    // 确定要显示的孔号
    let holeNumbersToShow;
    if (holeNumber === 'all') {
        // 显示所有孔号
        holeNumbersToShow = Object.keys(holeGroups);
    } else if (holeNumber) {
        // 显示指定孔号
        holeNumbersToShow = [holeNumber];
    } else {
        // 显示所有孔号（默认情况）
        holeNumbersToShow = Object.keys(holeGroups);
    }
    
    // 遍历每个孔号的数据
    holeNumbersToShow.forEach(holeNumber => {
        const holeData = holeGroups[holeNumber];
        if (!holeData) return;
        
        const totalHammerCount = holeData.length;
        
        // 找出深度最大值
        let maxDepth = 0;
        holeData.forEach(row => {
            const depth = parseFloat(row[depthColumn]);
            if (depth > maxDepth) {
                maxDepth = depth;
            }
        });
        
        // 找出深度最大值对应的锤击次数
        const significantHammers = holeData.filter(row => {
            const depth = parseFloat(row[depthColumn]);
            return depth === maxDepth;
        });
        
        // 生成孔号统计卡片
        resultHTML += `
        <div class="hole-statistic-card">
            <div class="hole-header">
                <span class="hole-number">孔号: <span class="number">${holeNumber}</span></span>
            </div>
            <div class="hole-stats">
                <p>总锤击数 <span class="number-primary">${totalHammerCount}</span> 总深度 <span class="number-primary">${maxDepth}</span> mm</p>
                ${significantHammers.map(row => {
                    const hammerCount = row[hammerCountColumn];
                    const depth = row[depthColumn];
                    const currentTotalDepth = row[totalDepthColumn];
                    return `<p>第 <span class="number-secondary">${hammerCount}</span> 次锤击 贯入 <span class="number-secondary">${depth}</span> mm/总深 <span class="number-secondary">${currentTotalDepth}</span> mm</p>`;
                }).join('')}
            </div>
        </div>
        `;
    });
    
    resultHTML += '</div>';
    calculateResult.innerHTML = resultHTML;
}

/**
 * 获取当前用于计算的数据
 * @returns {Array} 筛选后的数据
 */
function getCurrentDataForCalculate() {
    if (!currentHoleNumber || !parsedData || parsedData.length === 0) {
        return parsedData;
    }
    
    const headers = Object.keys(parsedData[0]);
    const holeNumberColumn = getColumnName(headers, COLUMN_INDEXES.HOLE_NUMBER);
    
    // 筛选当前孔号的数据
    return parsedData.filter(row => row[holeNumberColumn] === currentHoleNumber);
}

// 导出相关函数

// 导出图表为图片
function exportChartAsImage() {
    if (!chart) {
        showToast('请先导入数据并生成图表！', 'warning');
        return;
    }
    
    showLoading();
    html2canvas(chartCanvas).then(canvas => {
        const link = document.createElement('a');
        link.download = 'chart.png';
        link.href = canvas.toDataURL('image/png');
        link.click();
        hideLoading();
        showToast('图表导出为图片成功！', 'success');
    }).catch(error => {
        console.error('导出图片失败:', error);
        hideLoading();
        showToast('图表导出失败，请重试！', 'error');
    });
}

// 导出结果为Excel
function exportResultAsExcel() {
    if (!parsedData || parsedData.length === 0) {
        showToast('没有数据可导出！', 'warning');
        return;
    }
    
    showLoading();
    try {
        const worksheet = XLSX.utils.json_to_sheet(parsedData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '数据');
        XLSX.writeFile(workbook, 'result.xlsx');
        hideLoading();
        showToast('数据导出为Excel成功！', 'success');
    } catch (error) {
        console.error('导出Excel失败:', error);
        hideLoading();
        showToast('数据导出失败，请重试！', 'error');
    }
}

// 重置应用
function resetApp() {
    // 清空全局变量
    uploadedFile = null;
    parsedData = [];
    
    // 销毁图表
    if (chart) {
        chart.destroy();
        chart = null;
    }
    
    // 清空界面
    document.getElementById('dataTable').innerHTML = '';
    document.getElementById('calculateResult').innerHTML = '';
    document.getElementById('holeButtons').innerHTML = '';
    
    // 隐藏相关区域
    document.getElementById('previewSection').style.display = 'none';
    document.getElementById('calculateSection').style.display = 'none';
    document.getElementById('chartSection').style.display = 'none';
    document.getElementById('holeSelector').style.display = 'none';
    
    // 重置文件输入
    fileInput.value = '';
    
    console.log('应用已重置');
    showToast('应用已成功重置！', 'success');
}

// 启动应用
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
} else {
    initApp();
}
