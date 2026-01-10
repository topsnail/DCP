// CSV数据可视化应用
class CSVDataApp {
    constructor() {
        this.rawData = [];
        this.filteredData = [];
        this.charts = [];
        this.init();
    }

    // 初始化应用
    init() {
        this.bindEventListeners();
        this.loadCSVData();
        this.initBackToTop();
    }

    // Toast通知方法
    showToast(message, type = 'info') {
        const toastContainer = document.getElementById('toastContainer');
        if (!toastContainer) return;

        // 将换行符转换为HTML换行标签
        const formattedMessage = message.replace(/\n/g, '<br>');
        
        // 创建toast元素
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        
        // 根据类型设置图标
        const icons = {
            success: '✅',
            error: '❌',
            warning: '⚠️',
            info: 'ℹ️'
        };
        
        toast.innerHTML = `
            <span class="toast-icon">${icons[type] || icons.info}</span>
            <span class="toast-message">${formattedMessage}</span>
        `;
        
        // 添加到容器
        toastContainer.appendChild(toast);
        
        // 立即显示动画
        setTimeout(() => {
            toast.classList.add('show');
        }, 10);
        
        // 5秒后自动移除，给用户足够时间阅读长错误信息
        setTimeout(() => {
            toast.classList.remove('show');
            // 等待动画结束后移除元素
            setTimeout(() => {
                toast.remove();
            }, 300);
        }, 5000);
    }

    // 绑定事件监听器
    bindEventListeners() {
        // 文件导入事件
        const fileInput = document.getElementById('fileInput');
        if (fileInput) {
            fileInput.addEventListener('change', (e) => this.handleFileImport(e));
        }
        
        // 孔号筛选事件
        const holeNumberSelect = document.getElementById('holeNumber');
        if (holeNumberSelect) {
            holeNumberSelect.addEventListener('change', () => this.applyFilter());
        }
        
        // 导出图片事件
        const exportImageBtn = document.getElementById('exportImageBtn');
        if (exportImageBtn) {
            exportImageBtn.addEventListener('click', () => this.exportImage());
        }
        
        // 导出PDF事件
        const exportPdfBtn = document.getElementById('exportPdfBtn');
        if (exportPdfBtn) {
            exportPdfBtn.addEventListener('click', () => this.exportPDF());
        }
        
        // 清空重置事件
        const clearBtn = document.getElementById('clearBtn');
        if (clearBtn) {
            clearBtn.addEventListener('click', () => this.clearData());
        }
        
        // 返回顶部事件
        const backToTopBtn = document.getElementById('backToTopBtn');
        if (backToTopBtn) {
            backToTopBtn.addEventListener('click', () => this.scrollToTop());
        }
    }
    
    // 初始化返回顶部功能
    initBackToTop() {
        // 添加滚动事件监听器
        window.addEventListener('scroll', () => this.handleScroll());
        // 初始检测滚动位置
        this.handleScroll();
    }
    
    // 处理滚动事件，显示或隐藏返回顶部按钮
    handleScroll() {
        const backToTopBtn = document.getElementById('backToTopBtn');
        if (!backToTopBtn) return;
        
        // 当滚动超过300px时显示按钮
        if (window.scrollY > 300) {
            backToTopBtn.classList.add('show');
        } else {
            backToTopBtn.classList.remove('show');
        }
    }
    
    // 滚动到顶部
    scrollToTop() {
        // 使用平滑滚动效果
        window.scrollTo({
            top: 0,
            behavior: 'smooth'
        });
    }

    // 处理文件导入
    handleFileImport(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const text = e.target.result;
                this.rawData = this.parseCSV(text);
                this.initFilterOptions();
                this.applyFilter();
                this.showToast('文件导入成功！', 'success');
            } catch (error) {
                this.showToast(`文件解析失败: ${error.message}`, 'error');
            }
        };
        reader.readAsText(file);
    }

    // 加载默认CSV数据
    loadCSVData() {
        fetch('示例数据.CSV')
            .then(response => {
                if (!response.ok) {
                    throw new Error(`Network response was not ok: ${response.status}`);
                }
                return response.text();
            })
            .then(text => {
                this.rawData = this.parseCSV(text);
                this.initFilterOptions();
                this.applyFilter();
                this.showToast('默认CSV数据加载成功！', 'success');
            })
            .catch(error => {
                // 只在控制台输出错误，不向用户显示提示，因为这是正常情况
                console.log('未找到默认CSV文件，等待用户导入数据...');
            });
    }

    // 解析CSV数据
    parseCSV(csvText) {
        const lines = csvText.split('\n');
        const data = [];
        
        // 跳过表头行，直接从数据行开始解析
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (line) {
                const columns = line.split(',');
                
                // 只处理有7列的数据行
                if (columns.length >= 7) {
                    const date = columns[0].trim();
                    const time = columns[1].trim();
                    const holeNumber = parseInt(columns[2]) || 0;
                    const rodNumber = parseInt(columns[3]) || 0;
                    const blowCount = parseInt(columns[4]) || 0;
                    const depth = parseInt(columns[5]) || 0;
                    const totalDepth = parseInt(columns[6]) || 0;
                    
                    // 直接添加到数据中
                    data.push({
                        date: date,
                        time: time,
                        holeNumber: holeNumber,
                        rodNumber: rodNumber,
                        blowCount: blowCount,
                        depth: depth,
                        totalDepth: totalDepth
                    });
                }
            }
        }
        
        // 如果没有有效数据，返回空数组
        return data;
    }

    // 初始化筛选选项
    initFilterOptions() {
        // 提取唯一的孔号
        const holeNumbers = [...new Set(this.rawData.map(item => item.holeNumber))].sort((a, b) => a - b);

        // 填充孔号下拉菜单
        const holeNumberSelect = document.getElementById('holeNumber');
        if (!holeNumberSelect) return;
        
        // 清空现有选项
        holeNumberSelect.innerHTML = '<option value="">所有孔号</option>';
        
        // 添加孔号选项
        holeNumbers.forEach(id => {
            const option = document.createElement('option');
            option.value = id;
            option.textContent = id;
            holeNumberSelect.appendChild(option);
        });
    }

    // 应用筛选
    applyFilter() {
        const holeNumber = document.getElementById('holeNumber').value;
        
        // 根据孔号筛选数据
        if (!holeNumber) {
            this.filteredData = [...this.rawData];
        } else {
            this.filteredData = this.rawData.filter(item => 
                item.holeNumber === parseInt(holeNumber)
            );
        }
        
        // 更新图表、表格和统计信息
        this.updateCharts();
        this.updateTable();
        this.updateStats();
    }

    // 更新图表
    updateCharts() {
        // 清空现有图表
        this.charts.forEach(chart => chart.destroy());
        this.charts = [];
        
        const chartContainer = document.getElementById('chartContainer');
        if (!chartContainer) return;
        
        // 清空容器
        chartContainer.innerHTML = '';

        // 按孔号分组数据
        const dataByHole = this.groupDataByHole();
        
        // 为每个孔号创建图表
        for (const holeNumberStr in dataByHole) {
            const holeData = dataByHole[holeNumberStr];
            const holeNumber = parseInt(holeNumberStr);
            
            // 按锤击数排序数据
            holeData.sort((a, b) => a.blowCount - b.blowCount);
            
            // 准备图表数据
            const labels = holeData.map(item => item.blowCount);
            const depths = holeData.map(item => item.depth);
            const totalDepths = holeData.map(item => item.totalDepth);
            
            // 创建图表容器
            const chartItem = document.createElement('div');
            chartItem.className = 'chart-item';
            
            // 创建canvas元素
            const canvas = document.createElement('canvas');
            canvas.id = `chart-${holeNumber}`;
            chartItem.appendChild(canvas);
            
            // 添加到容器
            chartContainer.appendChild(chartItem);
            
            // 使用Chart.js创建图表
            const ctx = canvas.getContext('2d');
            const chart = new Chart(ctx, {
                type: 'line',
                data: {
                    labels: labels,
                    datasets: [
                        {
                            label: '深度 (单位: mm)',
                            data: depths,
                            borderColor: 'rgb(75, 192, 192)',
                            backgroundColor: 'rgba(75, 192, 192, 0.2)',
                            borderWidth: 2,
                            tension: 0.1,
                            yAxisID: 'y'
                        },
                        {
                            label: '总深度 (单位: mm)',
                            data: totalDepths,
                            borderColor: 'rgb(255, 99, 132)',
                            backgroundColor: 'rgba(255, 99, 132, 0.2)',
                            borderWidth: 2,
                            tension: 0.1,
                            yAxisID: 'y1'
                        }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        tooltip: {
                            enabled: true,
                            mode: 'index',
                            intersect: false,
                            backgroundColor: 'rgba(0, 0, 0, 0.8)',
                            titleColor: '#fff',
                            bodyColor: '#fff',
                            borderColor: '#fff',
                            borderWidth: 1,
                            padding: 12,
                            displayColors: true,
                            callbacks: {
                                title: function(context) {
                                    return `锤击数: ${context[0].label}`;
                                },
                                label: function(context) {
                                    return `${context.dataset.label}: ${context.parsed.y} mm`;
                                }
                            }
                        },
                        legend: {
                            position: 'top',
                        },
                        title: {
                        display: true,
                        text: `孔号 ${holeNumber} - 锤击数与深度关系图`,
                        font: {
                            size: 14
                        }
                    }
                    },
                    scales: {
                        y: {
                            type: 'linear',
                            display: true,
                            position: 'left',
                            title: {
                                display: true,
                                text: '深度 (mm)'
                            }
                        },
                        y1: {
                            type: 'linear',
                            display: true,
                            position: 'right',
                            title: {
                                display: true,
                                text: '总深度 (mm)'
                            },
                            grid: {
                                drawOnChartArea: false,
                            },
                        }
                    }
                }
            });
            
            this.charts.push(chart);
        }
    }

    // 按孔号分组数据
    groupDataByHole() {
        const groupedData = {};
        
        this.filteredData.forEach(item => {
            const holeNumber = item.holeNumber;
            if (!groupedData[holeNumber]) {
                groupedData[holeNumber] = [];
            }
            groupedData[holeNumber].push(item);
        });
        
        return groupedData;
    }

    // 更新表格
    updateTable() {
        const tableBody = document.getElementById('tableBody');
        if (!tableBody) return;
        
        // 计算深度和总深度的最大值，用于数据条显示
        const maxDepth = Math.max(...this.filteredData.map(item => item.depth), 1);
        const maxTotalDepth = Math.max(...this.filteredData.map(item => item.totalDepth), 1);
        
        // 构建HTML字符串，一次性插入DOM
        let html = '';
        this.filteredData.forEach(item => {
            // 计算深度数据条的宽度比例
            const depthPercentage = (item.depth / maxDepth) * 100;
            const totalDepthPercentage = (item.totalDepth / maxTotalDepth) * 100;
            
            html += `
                <tr>
                    <td>${item.date}</td>
                    <td>${item.time}</td>
                    <td>${item.holeNumber}</td>
                    <td>${item.rodNumber}</td>
                    <td>${item.blowCount}</td>
                    <td class="data-bar-cell">
                        <div class="data-bar-container">
                            <div class="data-bar depth-bar" style="width: ${depthPercentage}%"></div>
                            <span class="data-bar-value">${item.depth}</span>
                        </div>
                    </td>
                    <td class="data-bar-cell">
                        <div class="data-bar-container">
                            <div class="data-bar total-depth-bar" style="width: ${totalDepthPercentage}%"></div>
                            <span class="data-bar-value">${item.totalDepth}</span>
                        </div>
                    </td>
                </tr>
            `;
        });
        
        // 一次性插入所有行，减少DOM重排
        tableBody.innerHTML = html;
    }

    // 更新统计信息
    updateStats() {
        // 计算统计数据
        const totalRows = this.rawData.length;
        const holeCount = [...new Set(this.rawData.map(item => item.holeNumber))].length;
        const displayedRows = this.filteredData.length;
        
        // 更新DOM元素
        const totalRowsElement = document.getElementById('totalRows');
        if (totalRowsElement) totalRowsElement.textContent = totalRows;
        
        const holeCountElement = document.getElementById('holeCount');
        if (holeCountElement) holeCountElement.textContent = holeCount;
        
        const displayedRowsElement = document.getElementById('displayedRows');
        if (displayedRowsElement) displayedRowsElement.textContent = displayedRows;
        
        // 更新孔号深度统计
        this.updateHoleDepthStats();
    }

    // 更新孔号深度统计
    updateHoleDepthStats() {
        // 获取当前选择的孔号
        const selectedHoleNumber = document.getElementById('holeNumber').value;
        
        // 按孔号分组数据
        const dataByHole = {};
        this.rawData.forEach(item => {
            if (!dataByHole[item.holeNumber]) {
                dataByHole[item.holeNumber] = [];
            }
            dataByHole[item.holeNumber].push(item);
        });
        
        // 获取孔号深度统计容器
        const holeDepthStatsContainer = document.getElementById('holeDepthStats');
        if (!holeDepthStatsContainer) return;
        
        // 清空现有内容
        holeDepthStatsContainer.innerHTML = '';
        
        // 如果没有数据，显示提示
        if (Object.keys(dataByHole).length === 0) {
            holeDepthStatsContainer.innerHTML = '<div class="hole-stat-item">暂无数据</div>';
            return;
        }
        
        // 确定要显示的孔号列表
        let holeNumbersToShow = Object.keys(dataByHole).sort((a, b) => parseInt(a) - parseInt(b));
        
        // 如果选择了特定孔号，只显示该孔号
        if (selectedHoleNumber) {
            holeNumbersToShow = [selectedHoleNumber];
        }
        
        // 为每个孔号生成统计信息
        holeNumbersToShow.forEach(holeNumberStr => {
            const holeNumber = parseInt(holeNumberStr);
            const holeData = dataByHole[holeNumber];
            
            // 计算该孔号的最大深度
            const maxDepth = Math.max(...holeData.map(item => item.depth), 0);
            
            // 找出所有具有最大深度的记录
            const maxDepthRecords = holeData.filter(item => item.depth === maxDepth);
            
            // 计算该孔总锤击数
            const totalBlowCount = Math.max(...holeData.map(item => item.blowCount), 0);
            
            // 计算该孔总深度（最大总深度）
            const totalDepth = Math.max(...holeData.map(item => item.totalDepth), 0);
            
            // 创建孔号统计项
            const holeStatItem = document.createElement('div');
            holeStatItem.className = 'hole-stat-item';
            
            // 创建孔号标题
            const holeStatHeader = document.createElement('div');
            holeStatHeader.className = 'hole-stat-header';
            holeStatHeader.textContent = `孔号: ${holeNumber}`;
            holeStatItem.appendChild(holeStatHeader);
            
            // 显示总锤击数和总深度
            const totalInfo = document.createElement('div');
            totalInfo.className = 'hole-stat-detail';
            totalInfo.innerHTML = `总锤击数<span class="hole-stat-value">${totalBlowCount}</span>   总深度<span class="hole-stat-value">${totalDepth}</span>mm`;
            holeStatItem.appendChild(totalInfo);
            
            // 显示所有最大贯入深度记录
            maxDepthRecords.forEach((record, index) => {
                const maxDepthBlowCount = record.blowCount;
                const maxDepthTotalDepth = record.totalDepth;
                
                // 显示最大贯入深度信息
                const maxPenetrationInfo = document.createElement('div');
                maxPenetrationInfo.className = 'hole-stat-detail';
                maxPenetrationInfo.innerHTML = `第<span class="hole-stat-value">${maxDepthBlowCount}</span>次锤击 贯入<span class="hole-stat-value">${maxDepth}</span>mm/总深<span class="hole-stat-value">${maxDepthTotalDepth}</span>mm`;
                holeStatItem.appendChild(maxPenetrationInfo);
            });
            
            // 添加到容器
            holeDepthStatsContainer.appendChild(holeStatItem);
        });
    }

    // 导出图片
    async exportImage() {
        // 检查是否有图表
        if (this.charts.length === 0) {
            this.showToast('没有可导出的图表', 'warning');
            return;
        }
        
        try {
            // 获取图表容器
            const chartContainer = document.getElementById('chartContainer');
            if (!chartContainer) {
                throw new Error('未找到图表容器');
            }
            
            // 使用html2canvas生成图片
            const canvas = await html2canvas(chartContainer, {
                scale: 2, // 提高分辨率
                useCORS: true,
                logging: false
            });
            
            // 转换为图片链接并下载
            const link = document.createElement('a');
            link.download = `孔号曲线图_${new Date().getTime()}.png`;
            link.href = canvas.toDataURL('image/png');
            link.click();
            
            this.showToast('图片导出成功！', 'success');
        } catch (error) {
            this.showToast(`导出图片失败: ${error.message}`, 'error');
        }
    }

    // 导出PDF
    async exportPDF() {
        // 检查是否有图表
        if (this.charts.length === 0) {
            this.showToast('没有可导出的图表', 'warning');
            return;
        }
        
        try {
            // 导入jspdf
            const { jsPDF } = window.jspdf;
            const pdf = new jsPDF();
            
            // 获取图表容器
            const chartContainer = document.getElementById('chartContainer');
            if (!chartContainer) {
                throw new Error('未找到图表容器');
            }
            
            // 使用html2canvas生成图表图片
            const canvas = await html2canvas(chartContainer, {
                scale: 2, // 提高分辨率
                useCORS: true,
                logging: false
            });
            
            // 转换为图片数据
            const imgData = canvas.toDataURL('image/png');
            
            // 计算图片在PDF中的尺寸
            const imgWidth = 190; // PDF宽度，单位mm
            const imgHeight = (canvas.height * imgWidth) / canvas.width;
            
            // 添加图片到PDF
            pdf.addImage(imgData, 'PNG', 10, 10, imgWidth, imgHeight);
            
            // 保存PDF
            pdf.save(`孔号曲线图_${new Date().getTime()}.pdf`);
            
            this.showToast('PDF导出成功！', 'success');
        } catch (error) {
            this.showToast(`导出PDF失败: ${error.message}`, 'error');
        }
    }

    // 清空重置数据
    clearData() {
        // 清空数据
        this.rawData = [];
        this.filteredData = [];
        
        // 销毁所有图表
        this.charts.forEach(chart => chart.destroy());
        this.charts = [];
        
        // 清空图表容器
        const chartContainer = document.getElementById('chartContainer');
        if (chartContainer) {
            chartContainer.innerHTML = '';
        }
        
        // 清空表格
        const tableBody = document.getElementById('tableBody');
        if (tableBody) {
            tableBody.innerHTML = '';
        }
        
        // 重置孔号选择器
        const holeNumberSelect = document.getElementById('holeNumber');
        if (holeNumberSelect) {
            holeNumberSelect.innerHTML = '<option value="">所有孔号</option>';
        }
        
        // 重置文件输入
        const fileInput = document.getElementById('fileInput');
        if (fileInput) {
            fileInput.value = '';
        }
        
        // 更新统计信息
        this.updateStats();
        
        this.showToast('数据已清空重置！', 'success');
    }
}

// 页面加载完成后初始化应用
document.addEventListener('DOMContentLoaded', () => {
    try {
        new CSVDataApp();
    } catch (error) {
        console.error('初始化应用时发生错误:', error);
        // 使用简单的toast实现，因为此时CSVDataApp实例可能未创建
        const toastContainer = document.getElementById('toastContainer');
        if (toastContainer) {
            const toast = document.createElement('div');
            toast.className = 'toast error';
            toast.innerHTML = `
                <span class="toast-icon">❌</span>
                <span class="toast-message">初始化应用失败: ${error.message}</span>
            `;
            toastContainer.appendChild(toast);
            
            // 立即显示动画
            setTimeout(() => {
                toast.classList.add('show');
            }, 10);
            
            // 3秒后自动移除
            setTimeout(() => {
                toast.classList.remove('show');
                // 等待动画结束后移除元素
                setTimeout(() => {
                    toast.remove();
                }, 300);
            }, 3000);
        }
    }
});