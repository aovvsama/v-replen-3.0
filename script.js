/**
 * V-Replenishment Tool - 门店智能补货清单生成器
 * 版本: 1.0
 * 功能: 自动识别新品和缺货产品，生成专业补货清单
 */

// 全局变量
let inventoryData = [];
let arrivalData = [];
let replenishmentList = [];
let filteredList = [];

// DOM元素
const inventoryFileInput = document.getElementById('inventory-file');
const arrivalFileInput = document.getElementById('arrival-file');
const inventoryFilename = document.getElementById('inventory-filename');
const arrivalFilename = document.getElementById('arrival-filename');
const processBtn = document.getElementById('process-btn');
const exportBtn = document.getElementById('export-btn');
const printBtn = document.getElementById('print-btn');
const resultSection = document.querySelector('.result-section');
const resultTableBody = document.getElementById('result-table-body');
const tableSearch = document.getElementById('table-search');

// 统计元素
const totalItemsSpan = document.getElementById('total-items');
const newItemsSpan = document.getElementById('new-items');
const zeroStockItemsSpan = document.getElementById('zero-stock-items');
const newCCCountSpan = document.getElementById('new-cc-count');
const replenishSKUCountSpan = document.getElementById('replenish-sku-count');

// 模态框元素
const loadingOverlay = document.getElementById('loading');
const aboutModal = document.getElementById('about-modal');
const helpModal = document.getElementById('help-modal');

// 初始化函数
function init() {
    setupEventListeners();
    setupDragAndDrop();
    console.log('V-Replenishment Tool v1.0 已加载');
}

// 设置事件监听器
function setupEventListeners() {
    // 文件上传事件
    inventoryFileInput.addEventListener('change', handleFileUpload);
    arrivalFileInput.addEventListener('change', handleFileUpload);
    
    // 按钮事件
    processBtn.addEventListener('click', processData);
    exportBtn.addEventListener('click', exportToExcel);
    printBtn.addEventListener('click', printResults);
    
    // 搜索事件
    tableSearch.addEventListener('input', filterTable);
    
    // 模态框事件
    document.querySelectorAll('.close-modal').forEach(btn => {
        btn.addEventListener('click', closeAllModals);
    });
    
    document.querySelectorAll('.modal').forEach(modal => {
        modal.addEventListener('click', (e) => {
            if (e.target === modal) closeAllModals();
        });
    });
    
    // 键盘事件
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') closeAllModals();
    });
}

// 设置拖拽功能
function setupDragAndDrop() {
    setupDragDrop('inventory-box', inventoryFileInput);
    setupDragDrop('arrival-box', arrivalFileInput);
}

function setupDragDrop(elementId, fileInput) {
    const element = document.getElementById(elementId);
    
    element.addEventListener('dragover', (e) => {
        e.preventDefault();
        element.classList.add('drag-over');
    });
    
    element.addEventListener('dragleave', (e) => {
        e.preventDefault();
        element.classList.remove('drag-over');
    });
    
    element.addEventListener('drop', (e) => {
        e.preventDefault();
        element.classList.remove('drag-over');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            const file = files[0];
            if (isValidFileType(file)) {
                updateFileInput(fileInput, file);
                fileInput.dispatchEvent(new Event('change'));
            } else {
                alert('请上传Excel或CSV文件（.xlsx, .xls, .csv）');
            }
        }
    });
}

// 文件上传处理
async function handleFileUpload(event) {
    const fileInput = event.target;
    const file = fileInput.files[0];
    
    if (!file) return;
    
    // 更新文件名显示
    const filenameDisplay = fileInput.id === 'inventory-file' ? inventoryFilename : arrivalFilename;
    filenameDisplay.textContent = `📄 ${file.name} (${formatFileSize(file.size)})`;
    filenameDisplay.style.color = '#2d3748';
    
    try {
        showLoading(true);
        const data = await readFile(file);
        
        if (fileInput.id === 'inventory-file') {
            inventoryData = data;
            console.log('库存数据加载:', inventoryData.length, '行');
        } else {
            arrivalData = data;
            console.log('到货数据加载:', arrivalData.length, '行');
        }
        
        // 验证数据格式
        if (!validateDataFormat(fileInput.id, data)) {
            filenameDisplay.innerHTML = `⚠️ ${file.name} - 请检查数据格式`;
            filenameDisplay.style.color = '#e53e3e';
            
            if (fileInput.id === 'inventory-file') {
                inventoryData = [];
            } else {
                arrivalData = [];
            }
        }
        
        checkFilesReady();
        
    } catch (error) {
        console.error('文件读取错误:', error);
        const filenameDisplay = fileInput.id === 'inventory-file' ? inventoryFilename : arrivalFilename;
        filenameDisplay.innerHTML = `❌ ${file.name} - 读取失败`;
        filenameDisplay.style.color = '#e53e3e';
        alert('文件读取失败，请检查文件格式是否正确。');
    } finally {
        showLoading(false);
    }
}

// 读取文件
async function readFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = e.target.result;
                let result;
                
                if (file.name.toLowerCase().endsWith('.csv')) {
                    result = parseCSV(data);
                } else {
                    const workbook = XLSX.read(data, { 
                        type: file.name.toLowerCase().endsWith('.xlsx') ? 'array' : 'binary',
                        raw: true,
                        cellDates: true,
                        cellStyles: true
                    });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    result = XLSX.utils.sheet_to_json(firstSheet, { 
                        defval: '',
                        raw: false,
                        dateNF: 'yyyy-mm-dd'
                    });
                }
                
                resolve(result);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = reject;
        
        if (file.name.toLowerCase().endsWith('.csv')) {
            reader.readAsText(file, 'UTF-8');
        } else if (file.name.toLowerCase().endsWith('.xlsx')) {
            reader.readAsArrayBuffer(file);
        } else {
            reader.readAsBinaryString(file);
        }
    });
}

// 解析CSV数据
function parseCSV(csvText) {
    const lines = csvText.split(/\r\n|\n/).map(line => line.trim()).filter(line => line);
    if (lines.length === 0) return [];
    
    // 检测分隔符
    const delimiter = detectDelimiter(lines[0]);
    
    const headers = lines[0].split(delimiter).map(h => h.trim().replace(/^"|"$/g, ''));
    const result = [];
    
    for (let i = 1; i < lines.length; i++) {
        const values = parseCSVLine(lines[i], delimiter);
        const row = {};
        
        headers.forEach((header, index) => {
            row[header] = values[index] || '';
        });
        
        result.push(row);
    }
    
    return result;
}

// 检测CSV分隔符
function detectDelimiter(line) {
    const commaCount = (line.match(/,/g) || []).length;
    const semicolonCount = (line.match(/;/g) || []).length;
    const tabCount = (line.match(/\t/g) || []).length;
    
    if (tabCount > commaCount && tabCount > semicolonCount) return '\t';
    if (semicolonCount > commaCount) return ';';
    return ',';
}

// 解析CSV行（处理引号内的分隔符）
function parseCSVLine(line, delimiter) {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        const nextChar = line[i + 1];
        
        if (char === '"') {
            if (inQuotes && nextChar === '"') {
                current += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (char === delimiter && !inQuotes) {
            result.push(current.trim());
            current = '';
        } else {
            current += char;
        }
    }
    
    result.push(current.trim());
    return result;
}

// 验证数据格式
function validateDataFormat(type, data) {
    if (!data || data.length === 0) return false;
    
    const firstRow = data[0];
    const keys = Object.keys(firstRow);
    
    if (type === 'inventory-file') {
        // 检查库存表必需字段
        const requiredFields = ['规格编码', '商品名称', '总库存'];
        const hasRequired = requiredFields.some(field => 
            keys.some(key => key.includes(field))
        );
        return hasRequired;
    } else {
        // 检查到货表必需字段
        const requiredFields = ['Barcode', 'Item Number', 'Product name', 'Order Qty'];
        const hasRequired = requiredFields.some(field => 
            keys.some(key => key.includes(field))
        );
        return hasRequired;
    }
}

// 检查文件是否就绪
function checkFilesReady() {
    const isReady = inventoryData.length > 0 && arrivalData.length > 0;
    processBtn.disabled = !isReady;
    
    if (isReady) {
        processBtn.innerHTML = '<i class="fas fa-cogs"></i> 智能分析并生成补货清单';
    } else {
        processBtn.innerHTML = '<i class="fas fa-cogs"></i> 请上传两个文件';
    }
}

// 主要处理函数
async function processData() {
    showLoading(true);
    
    try {
        await new Promise(resolve => setTimeout(resolve, 100));
        
        // 1. 创建库存映射
        const inventoryMap = createInventoryMap();
        
        // 2. 处理到货数据
        replenishmentList = processArrivalData(inventoryMap);
        
        // 3. 排序：新品优先，然后按库存从低到高
        replenishmentList.sort((a, b) => {
            if (a.isNew && !b.isNew) return -1;
            if (!a.isNew && b.isNew) return 1;
            if (a.currentStock === 0 && b.currentStock !== 0) return -1;
            if (a.currentStock !== 0 && b.currentStock === 0) return 1;
            return a.currentStock - b.currentStock;
        });
        
        // 4. 更新UI
        filteredList = [...replenishmentList];
        updateResultsTable();
        updateStatistics();
        
        // 5. 启用导出按钮
        exportBtn.disabled = false;
        resultSection.classList.add('active');
        
        // 滚动到结果区域
        resultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
        
    } catch (error) {
        console.error('数据处理错误:', error);
        alert('数据处理失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// 创建库存映射
function createInventoryMap() {
    const map = new Map();
    const productInfoMap = new Map();
    
    inventoryData.forEach(item => {
        const sku = getFieldValue(item, ['规格编码', 'SKU', 'Barcode', '商品编码']);
        if (!sku) return;
        
        const stock = parseFloat(getFieldValue(item, ['总库存', '可用库存', '库存数', '库存数量']) || 0);
        const productName = getFieldValue(item, ['商品名称', '产品名称', 'Product name']);
        const brand = getFieldValue(item, ['商品品牌', '品牌', 'Brand']);
        const category = getFieldValue(item, ['商品分类', '分类', 'Category']);
        const retailPrice = getFieldValue(item, ['零售价', '价格', 'Price', 'MSRP']);
        
        map.set(sku.toString().trim(), stock);
        
        productInfoMap.set(sku.toString().trim(), {
            productName,
            brand,
            category,
            retailPrice
        });
    });
    
    console.log('库存映射创建完成，共', map.size, '个SKU');
    return { stockMap: map, productInfoMap };
}

// 处理到货数据
function processArrivalData(inventoryMap) {
    const result = [];
    let newProductCount = 0;
    let zeroStockCount = 0;
    
    arrivalData.forEach(item => {
        const barcode = getFieldValue(item, ['Barcode', 'SKU', '规格编码', '条码']);
        if (!barcode) return;
        
        const orderQty = parseInt(getFieldValue(item, ['Order Qty', '到货数量', '数量', 'Qty']) || 0);
        if (orderQty <= 0) return;
        
        const sku = barcode.toString().trim();
        const currentStock = inventoryMap.stockMap.get(sku) || 0;
        const isNew = !inventoryMap.stockMap.has(sku);
        
        // 只添加需要补货的产品（新品或库存为0）
        if (isNew || currentStock === 0) {
            if (isNew) newProductCount++;
            if (currentStock === 0) zeroStockCount++;
            
            // 获取产品信息
            const productInfo = inventoryMap.productInfoMap.get(sku) || {};
            
            const replenishmentItem = {
                barcode: sku,
                status: isNew ? 'NEW' : '补货',
                productGender: getFieldValue(item, ['ProductGender', '性别', 'Gender']) || '',
                itemNumber: getFieldValue(item, ['Item Number', '款号', '货号']) || '',
                productName: getFieldValue(item, ['Product name', '商品名称', '产品名称']) || productInfo.productName || '',
                color: getFieldValue(item, ['Color', '颜色', 'colour']) || '',
                size: getFieldValue(item, ['Size', '尺码', '规格']) || '',
                orderQty: orderQty,
                currentStock: currentStock,
                brand: productInfo.brand || '',
                category: productInfo.category || '',
                retailPrice: productInfo.retailPrice || '',
                isNew: isNew,
                isZeroStock: currentStock === 0,
                priority: calculatePriority(isNew, currentStock),
                batchNumber: getFieldValue(item, ['Order #', 'Order No', 'Batch', '批次']) || ''
            };
            
            result.push(replenishmentItem);
        }
    });
    
    console.log('补货清单生成完成:', {
        total: result.length,
        newProducts: newProductCount,
        zeroStock: zeroStockCount
    });
    
    return result;
}

// 计算优先级
function calculatePriority(isNew, currentStock) {
    if (isNew) return '高';
    if (currentStock === 0) return '中';
    return '低';
}

// 更新结果表格
function updateResultsTable() {
    if (filteredList.length === 0) {
        resultTableBody.innerHTML = `
            <tr class="empty-row">
                <td colspan="10">
                    <div class="empty-state">
                        <i class="fas fa-search"></i>
                        <p>未找到匹配的产品</p>
                    </div>
                </td>
            </tr>
        `;
        return;
    }
    
    let html = '';
    
    filteredList.forEach((item, index) => {
        const rowClass = item.isNew ? 'new-item' : 
                        item.isZeroStock ? 'zero-stock-item' : 
                        index % 2 === 0 ? '' : 'alternate-row';
        
        html += `
        <tr class="${rowClass}">
            <td>
                <span class="status-badge ${item.isNew ? 'status-new' : 'status-replenish'}">
                    ${item.isNew ? 'NEW' : '补货'}
                </span>
            </td>
            <td>${escapeHtml(item.productGender)}</td>
            <td><strong>${escapeHtml(item.itemNumber)}</strong></td>
            <td>${escapeHtml(item.productName)}</td>
            <td class="center-cell">${escapeHtml(item.color)}</td>
            <td class="center-cell">${escapeHtml(item.size)}</td>
            <td class="number-cell"><strong>${item.orderQty}</strong></td>
            <td class="number-cell ${item.currentStock === 0 ? 'stock-zero' : 'stock-available'}">
                ${item.currentStock}
            </td>
            <td>${escapeHtml(item.category)}</td>
            <td class="number-cell">${escapeHtml(formatPrice(item.retailPrice))}</td>
        </tr>
        `;
    });
    
    resultTableBody.innerHTML = html;
}

// 更新统计信息
function updateStatistics() {
    // 1. 总件数：到货表中所有 Order Qty 的总和
    const totalQty = arrivalData.reduce((sum, item) => {
        const qty = parseInt(getFieldValue(item, ['Order Qty', '到货数量', '数量', 'Qty']) || 0);
        return sum + (isNaN(qty) ? 0 : qty);
    }, 0);
    
    // 2. 新品总件数：新品的 Order Qty 总和
    const newTotalQty = replenishmentList
        .filter(item => item.isNew)
        .reduce((sum, item) => sum + item.orderQty, 0);
    
    // 3. 补货总件数：非新品的补货产品 Order Qty 总和
    const replenishTotalQty = replenishmentList
        .filter(item => !item.isNew)
        .reduce((sum, item) => sum + item.orderQty, 0);
    
    // 4. 新CC数量：新品按 Item Number + Color 组合去重计数
    const newCCSet = new Set();
    replenishmentList
        .filter(item => item.isNew)
        .forEach(item => {
            const ccKey = `${item.itemNumber}|${item.color}`;
            newCCSet.add(ccKey);
        });
    const newCCCount = newCCSet.size;
    
    // 5. 补货SKU数量：库存为0的产品按 Barcode 去重计数
    const zeroStockSKUSet = new Set();
    replenishmentList
        .filter(item => item.isZeroStock)
        .forEach(item => {
            zeroStockSKUSet.add(item.barcode);
        });
    const zeroStockSKUCount = zeroStockSKUSet.size;
    
    // 更新DOM
    totalItemsSpan.textContent = totalQty;
    newItemsSpan.textContent = newTotalQty;
    zeroStockItemsSpan.textContent = replenishTotalQty;
    newCCCountSpan.textContent = newCCCount;
    replenishSKUCountSpan.textContent = zeroStockSKUCount;
}

// 导出到Excel
async function exportToExcel() {
    if (replenishmentList.length === 0) {
        alert('没有数据可导出');
        return;
    }
    
    showLoading(true);
    
    try {
        // 创建工作簿
        const wb = XLSX.utils.book_new();
        
        // 获取批次信息
        const batchNumber = replenishmentList[0]?.batchNumber || '';
        
        // 计算总件数
        const totalQty = arrivalData.reduce((sum, item) => {
            const orderQty = parseInt(getFieldValue(item, ['Order Qty', '到货数量', '数量', 'Qty']) || 0);
            return sum + (isNaN(orderQty) ? 0 : orderQty);
        }, 0);
        
        // 准备数据
        const wsData = [
            ['V-Replenishment 智能补货清单', '', '', '', '', '', '', '', ''],
            [`生成时间: ${new Date().toLocaleString('zh-CN')}`, '', '', '', '', '', '', '', ''],
            [`门店: 北京三里屯`, '', '', '', '', '', '', '', ''],
            [`批次: ${batchNumber}`, '', '', '', '', '', '', '', ''],
            [`总件数: ${totalQty}`, '', '', '', '', '', '', '', ''],
            [],
            [
                '状态',
                'ProductGender',
                'Color Choice',
                'Product name',
                'Size',
                'Order Qty',
                '当前库存',
                '备注'
            ]
        ];
        
        // 添加数据行
        replenishmentList.forEach((item) => {
            const colorChoice = `${item.itemNumber}${item.color}`.trim();
            
            const row = [
                item.isNew ? 'NEW' : '补货',
                item.productGender,
                colorChoice,
                item.productName,
                item.size,
                item.orderQty,
                item.currentStock,
                item.isNew ? '新品首次到货' : (item.currentStock === 0 ? '库存为0需优先补货' : '库存不足需补货')
            ];
            
            wsData.push(row);
        });
        
        // 创建工作表
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        
        // 设置列宽
        const wscols = [
            { wch: 8 },
            { wch: 12 },
            { wch: 25 },
            { wch: 35 },
            { wch: 8 },
            { wch: 12 },
            { wch: 12 },
            { wch: 25 }
        ];
        ws['!cols'] = wscols;
        
        // 合并标题单元格
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } },
            { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
            { s: { r: 2, c: 0 }, e: { r: 2, c: 4 } },
            { s: { r: 3, c: 0 }, e: { r: 3, c: 4 } },
            { s: { r: 4, c: 0 }, e: { r: 4, c: 4 } }
        ];
        
        // 添加到工作簿
        XLSX.utils.book_append_sheet(wb, ws, '补货清单');
        
        // 生成文件名
        const date = new Date();
        const dateStr = date.toISOString().split('T')[0];
        const timeStr = date.getHours().toString().padStart(2, '0') + 
                       date.getMinutes().toString().padStart(2, '0');
        const fileName = `V-Replenishment_补货清单_${dateStr}_${timeStr}.xlsx`;
        
        // 导出文件
        XLSX.writeFile(wb, fileName);
        
        setTimeout(() => {
            alert('✅ Excel文件导出成功！\n文件已保存到您的下载文件夹。');
        }, 500);
        
    } catch (error) {
        console.error('导出错误:', error);
        alert('导出失败: ' + error.message);
    } finally {
        showLoading(false);
    }
}

// 打印结果
function printResults() {
    if (replenishmentList.length === 0) {
        alert('没有数据可打印');
        return;
    }
    
    const printWindow = window.open('', '_blank');
    printWindow.document.write(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>V-Replenishment 补货清单 - 打印</title>
            <meta charset="UTF-8">
            <style>
                @media print {
                    @page {
                        margin: 1cm;
                        size: A4 landscape;
                    }
                }
                body {
                    font-family: "等线", "Microsoft YaHei", sans-serif;
                    margin: 0;
                    padding: 20px;
                    color: #333;
                }
                .print-header {
                    text-align: center;
                    margin-bottom: 30px;
                    border-bottom: 3px solid #2c5282;
                    padding-bottom: 20px;
                }
                .print-title {
                    font-size: 24px;
                    font-weight: bold;
                    color: #2c5282;
                    margin-bottom: 10px;
                }
                .print-info {
                    font-size: 14px;
                    color: #666;
                    margin-bottom: 5px;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                    font-size: 12px;
                    margin-top: 20px;
                }
                th {
                    background-color: #2c5282;
                    color: white;
                    font-weight: bold;
                    padding: 12px 8px;
                    text-align: left;
                    border: 1px solid #ddd;
                }
                td {
                    padding: 10px 8px;
                    border: 1px solid #ddd;
                }
                tr:nth-child(even) {
                    background-color: #f8f9fa;
                }
                .new-row {
                    background-color: #fef3c7 !important;
                }
                .status-new {
                    background-color: #ed8936;
                    color: white;
                    padding: 4px 8px;
                    border-radius: 3px;
                    font-size: 11px;
                }
                .status-replenish {
                    background-color: #4299e1;
                    color: white;
                    padding: 4px 8px;
                    border-radius: 3px;
                    font-size: 11px;
                }
                .number {
                    text-align: right;
                }
                .stock-zero {
                    color: #e53e3e;
                    font-weight: bold;
                }
                .print-footer {
                    margin-top: 40px;
                    padding-top: 20px;
                    border-top: 1px solid #ddd;
                    font-size: 12px;
                    color: #666;
                    text-align: center;
                }
            </style>
        </head>
        <body>
            <div class="print-header">
                <div class="print-title">V-Replenishment 智能补货清单</div>
                <div class="print-info">生成时间: ${new Date().toLocaleString('zh-CN')}</div>
                <div class="print-info">门店: 北京三里屯 | 总计: ${replenishmentList.length} 个产品</div>
                <div class="print-info">新品: ${replenishmentList.filter(item => item.isNew).length} 个 | 缺货: ${replenishmentList.filter(item => item.isZeroStock).length} 个</div>
            </div>
            
            <table>
                <thead>
                    <tr>
                        <th>状态</th>
                        <th>ProductGender</th>
                        <th>Item Number</th>
                        <th>Product name</th>
                        <th>Color</th>
                        <th>Size</th>
                        <th>Order Qty</th>
                        <th>当前库存</th>
                        <th>分类</th>
                        <th>零售价</th>
                    </tr>
                </thead>
                <tbody>
                    ${replenishmentList.map(item => `
                        <tr class="${item.isNew ? 'new-row' : ''}">
                            <td>
                                <span class="${item.isNew ? 'status-new' : 'status-replenish'}">
                                    ${item.isNew ? 'NEW' : '补货'}
                                </span>
                            </td>
                            <td>${escapeHtml(item.productGender)}</td>
                            <td><strong>${escapeHtml(item.itemNumber)}</strong></td>
                            <td>${escapeHtml(item.productName)}</td>
                            <td>${escapeHtml(item.color)}</td>
                            <td>${escapeHtml(item.size)}</td>
                            <td class="number"><strong>${item.orderQty}</strong></td>
                            <td class="number ${item.currentStock === 0 ? 'stock-zero' : ''}">
                                ${item.currentStock}
                            </td>
                            <td>${escapeHtml(item.category)}</td>
                            <td class="number">${escapeHtml(formatPrice(item.retailPrice))}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
            
            <div class="print-footer">
                第 1 页 / 共 1 页 | V-Replenishment Tool v1.0 | 系统自动生成
            </div>
            
            <script>
                window.onload = function() {
                    window.print();
                    setTimeout(function() {
                        window.close();
                    }, 500);
                };
            </script>
        </body>
        </html>
    `);
    printWindow.document.close();
}

// 表格搜索过滤
function filterTable() {
    const searchTerm = tableSearch.value.toLowerCase().trim();
    
    if (!searchTerm) {
        filteredList = [...replenishmentList];
    } else {
        filteredList = replenishmentList.filter(item => {
            return (
                item.itemNumber.toLowerCase().includes(searchTerm) ||
                item.productName.toLowerCase().includes(searchTerm) ||
                item.color.toLowerCase().includes(searchTerm) ||
                (item.category && item.category.toLowerCase().includes(searchTerm))
            );
        });
    }
    
    updateResultsTable();
    updateStatistics();
}

// 工具函数
function isValidFileType(file) {
    const validTypes = ['.xlsx', '.xls', '.csv'];
    return validTypes.some(type => file.name.toLowerCase().endsWith(type));
}

function updateFileInput(fileInput, file) {
    const dataTransfer = new DataTransfer();
    dataTransfer.items.add(file);
    fileInput.files = dataTransfer.files;
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function getFieldValue(obj, possibleKeys) {
    for (const key of possibleKeys) {
        if (obj[key] !== undefined && obj[key] !== null && obj[key] !== '') {
            return obj[key];
        }
    }
    return '';
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text || '';
    return div.innerHTML;
}

function formatPrice(price) {
    if (!price) return '';
    const num = parseFloat(price.toString().replace(/[^\d.-]/g, ''));
    return isNaN(num) ? price : '¥' + num.toFixed(2);
}

function showLoading(show) {
    if (show) {
        loadingOverlay.classList.add('active');
        processBtn.disabled = true;
        processBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 处理中...';
    } else {
        loadingOverlay.classList.remove('active');
        checkFilesReady();
    }
}

// 模态框函数
function showAbout() {
    aboutModal.classList.add('active');
}

function showHelp() {
    helpModal.classList.add('active');
}

function closeAllModals() {
    document.querySelectorAll('.modal').forEach(modal => {
        modal.classList.remove('active');
    });
}

// 初始化应用
document.addEventListener('DOMContentLoaded', init);