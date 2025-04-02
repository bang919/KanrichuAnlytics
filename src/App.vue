<script setup>
import { ref, onMounted, watch } from 'vue';
import { ElMessage, ElMessageBox } from 'element-plus';
import * as XLSX from 'xlsx';
import * as ExcelJS from 'exceljs';
import FileSaver from 'file-saver';
// 导入HyperFormula库，用于计算Excel公式
import { HyperFormula } from 'hyperformula';
import { 
  Box, 
  Document, 
  DataAnalysis, 
  DataLine, 
  Upload, 
  Check,
  Delete,
  Download
} from '@element-plus/icons-vue';

// 文件上传区域引用
const dropZone = ref(null);
const fileList = ref([]);

// 上传状态
const uploading = ref(false);
const uploadStep = ref(1);
const fileReady = ref({
  xiyue: false,
  fba: false,
  sevenDay: false,
  thirtyDay: false
});

// 储存基础文件数据
const baseFileData = ref(null);
const resultFileData = ref(null);
const processedData = ref({
  inventoryMap: null,
  fbaInventoryMap: null,
  sevenDaySalesMap: null,
  thirtyDaySalesMap: null,
  unmatchedAsins: [], // 存储喜悦库存文件有但产品库存表中没有的ASIN
  statistics: {
    totalFactoryUpdated: 0,
    totalFBAUpdated: 0,
    totalSevenDayUpdated: 0,
    totalThirtyDayUpdated: 0
  }
});

// 在文件的data部分添加新的ref
const processedTemplates = ref({
  replenishmentTemplate: null,
  shippingTemplate: null,
  backendShippingTemplate: null
});

// 自动判断文件类型函数
const identifyFileType = (file) => {
  const fileName = file.name.toLowerCase();
  
  // 喜悦库存文件判断 - 名字包含"喜悦"的csv
  if (fileName.includes('喜悦') && fileName.endsWith('.csv')) {
    return 'xiyue';
  }
  
  // FBA库存文件判断 - "FBAInventory"开头的xlsx
  if (fileName.startsWith('fbainventory') && (fileName.endsWith('.xlsx') || fileName.endsWith('.xls'))) {
    return 'fba';
  }
  
  // 7天和30天产品分析 - 包含日期范围的判断
  const dateRangeMatch = fileName.match(/(\d{8})-(\d{8})/);
  if (dateRangeMatch) {
    const startDate = new Date(
      dateRangeMatch[1].substring(0, 4) + '-' + 
      dateRangeMatch[1].substring(4, 6) + '-' + 
      dateRangeMatch[1].substring(6, 8)
    );
    const endDate = new Date(
      dateRangeMatch[2].substring(0, 4) + '-' + 
      dateRangeMatch[2].substring(4, 6) + '-' + 
      dateRangeMatch[2].substring(6, 8)
    );
    
    const daysDiff = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24));
    
    if (daysDiff > 10) {
      return 'thirtyDay';
    } else {
      return 'sevenDay';
    }
  }
  
  // 无法识别
  return 'unknown';
};

// 处理文件上传
const handleFileUpload = (files) => {
  if (!files || files.length === 0) return;
  
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const fileType = identifyFileType(file);
    
    if (fileType === 'unknown') {
      ElMessage({
        message: `无法识别文件类型: ${file.name}`,
        type: 'warning'
      });
      continue;
    }
    
    // 添加到文件列表，带状态
    const newFile = {
      name: file.name,
      type: fileType,
      size: (file.size / 1024).toFixed(2) + ' KB',
      file: file,
      status: 'uploaded'
    };
    
    // 更新文件列表
    fileList.value = [...fileList.value.filter(f => f.type !== fileType), newFile];
    
    // 更新准备状态
    fileReady.value[fileType] = true;
    
    // 显示成功消息
    ElMessage({
      message: `成功上传 ${file.name} (${getFileTypeLabel(fileType)})`,
      type: 'success'
    });
  }
};

// 获取文件类型的中文标签
const getFileTypeLabel = (type) => {
  switch(type) {
    case 'xiyue': return '喜悦库存';
    case 'fba': return 'FBA库存';
    case 'sevenDay': return '7天产品分析';
    case 'thirtyDay': return '30天产品分析';
    default: return '未知类型';
  }
};

// 初始化拖放区域
onMounted(() => {
  if (dropZone.value) {
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
      dropZone.value.addEventListener(eventName, preventDefaults, false);
    });
    
    ['dragenter', 'dragover'].forEach(eventName => {
      dropZone.value.addEventListener(eventName, highlight, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
      dropZone.value.addEventListener(eventName, unhighlight, false);
    });
    
    dropZone.value.addEventListener('drop', handleDrop, false);
  }
});

// 阻止默认行为
function preventDefaults(e) {
  e.preventDefault();
  e.stopPropagation();
}

// 高亮拖放区
function highlight() {
  dropZone.value.classList.add('highlight');
}

// 取消高亮
function unhighlight() {
  dropZone.value.classList.remove('highlight');
}

// 处理拖放
function handleDrop(e) {
  const dt = e.dataTransfer;
  const files = dt.files;
  handleFileUpload(files);
}

// 通过点击触发文件选择
const triggerFileSelect = () => {
  const fileInput = document.createElement('input');
  fileInput.type = 'file';
  fileInput.multiple = true;
  fileInput.accept = '.csv,.xlsx,.xls';
  fileInput.onchange = (e) => handleFileUpload(e.target.files);
  fileInput.click();
};

// 删除文件
const removeFile = (file) => {
  fileList.value = fileList.value.filter(f => f !== file);
  fileReady.value[file.type] = false;
};

// 重置所有上传
const resetUploads = () => {
  ElMessageBox.confirm('确定要重置所有上传的文件吗?', '提示', {
    confirmButtonText: '确定',
    cancelButtonText: '取消',
    type: 'warning'
  }).then(() => {
    fileList.value = [];
    fileReady.value = {
      xiyue: false,
      fba: false,
      sevenDay: false,
      thirtyDay: false
    };
    uploadStep.value = 1;
    
    ElMessage({
      message: '已重置所有文件',
      type: 'success'
    });
  }).catch(() => {});
};

// 检查所有文件是否准备就绪
const allFilesReady = () => {
  return fileReady.value.xiyue && 
         fileReady.value.fba && 
         fileReady.value.sevenDay && 
         fileReady.value.thirtyDay;
};

// 处理文件函数
const processFiles = () => {
  if (!allFilesReady()) {
    ElMessage({
      message: '请先上传所有必需的文件',
      type: 'warning'
    });
    return;
  }
  
  uploading.value = true;
  
  // 获取上传的文件
  const xiyueFile = fileList.value.find(f => f.type === 'xiyue').file;
  const fbaFile = fileList.value.find(f => f.type === 'fba').file;
  const sevenDayFile = fileList.value.find(f => f.type === 'sevenDay').file;
  const thirtyDayFile = fileList.value.find(f => f.type === 'thirtyDay').file;
  
  // 加载基础文件
  loadBaseFile().then(() => {
    // 并行处理所有文件
    Promise.all([
      processXiyueInventory(xiyueFile),
      processFBAInventory(fbaFile),
      processSevenDayAnalysis(sevenDayFile),
      processThirtyDayAnalysis(thirtyDayFile)
    ]).then(() => {
      // 生成结果文件
      generateResultFile();
      
      // 显示成功消息
      ElMessage({
        message: '文件处理完成！',
        type: 'success'
      });
      
      uploading.value = false;
      uploadStep.value = 2; // 处理完成后进入第二步：查看处理后的库存表
      
      // 计算货值统计
      console.log('在processFiles中准备调用calculateValues');
  setTimeout(() => {
        calculateValues();
        console.log('在processFiles中调用了calculateValues');
      }, 800);
    }).catch(error => {
    uploading.value = false;
    ElMessage({
        message: `处理文件时出错: ${error.message}`,
        type: 'error'
      });
    });
  }).catch(error => {
    uploading.value = false;
    ElMessage({
      message: `加载基础文件时出错: ${error.message}`,
      type: 'error'
    });
  });
};

// 加载基础文件
const loadBaseFile = async () => {
  return new Promise((resolve, reject) => {
    try {
      // 创建一个请求，获取基础文件
      fetch('/产品库存及周转统计.xlsx')
        .then(response => {
          if (!response.ok) {
            throw new Error(`无法加载基础文件，状态码: ${response.status}`);
          }
          return response.arrayBuffer();
        })
        .then(arrayBuffer => {
          const data = new Uint8Array(arrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          
          // 将工作簿存储在ref中
          baseFileData.value = workbook;
          resolve();
        })
        .catch(error => {
          console.error("加载基础文件失败:", error);
          reject(error);
        });
    } catch (error) {
      console.error("加载基础文件异常:", error);
      reject(error);
    }
  });
};

// 处理喜悦库存文件
const processXiyueInventory = async (file) => {
  return new Promise((resolve, reject) => {
    try {
      const reader = new FileReader();
      
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          let workbook;
          
          // 尝试读取为CSV
          if (file.name.toLowerCase().endsWith('.csv')) {
            const csvContent = new TextDecoder('utf-8').decode(data);
            workbook = XLSX.read(csvContent, { type: 'string' });
          } else {
            // 否则读取为Excel
            workbook = XLSX.read(data, { type: 'array' });
          }
          
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          
          // 根据截图的确认，ASIN在第3行，剩余库存在第5行
          // 索引从0开始，所以实际为索引2和4
          const asinRow = 2; // 第3行
          const nameRow = 1;  // 第2行，通常是产品名称行
          const inventoryRow = 4; // 第5行
          
          if (!jsonData[asinRow] || !jsonData[inventoryRow]) {
            reject(new Error('文件格式不符合预期，缺少ASIN行或剩余库存行'));
            return;
          }
          
          // 创建ASIN到库存的映射
          const inventoryMap = {};
          // 创建ASIN到名称的映射，用于记录未匹配产品
          const nameMap = {};
          let foundCount = 0;
          
          // 找到有效数据的最后一列
          let lastValidColumn = 1;
          for (let col = jsonData[nameRow].length - 1; col >= 1; col--) {
            // 检查产品名称行，判断是否为有效产品名称
            const name = jsonData[nameRow][col];
            if (name && typeof name === 'string' && name.trim() !== '') {
              // 排除像"橡皮筋"、"挂钩"这样的辅助文本
              if (!['橡皮筋', '挂钩', '辅料', '配件'].includes(name.trim())) {
                lastValidColumn = col;
                break;
              }
            }
          }
          
          console.log(`找到有效数据的最后一列: ${lastValidColumn}`);
          
          // 从第1列开始处理（索引0是"在·"列，不含ASIN）
          for (let col = 1; col <= lastValidColumn; col++) {
            const asin = jsonData[asinRow][col];
            const name = jsonData[nameRow] && jsonData[nameRow][col];
            const inventoryCell = jsonData[inventoryRow][col];
            
            // 如果产品名称有值（即使ASIN没有）
            if (name && typeof name === 'string' && name.trim() !== '') {
              const nameStr = name.trim();
              
              // 尝试将库存转换为数字
              let inventory = Number(inventoryCell);
              if (isNaN(inventory)) {
                const matches = inventoryCell ? String(inventoryCell).match(/\d+/) : null;
                inventory = matches ? Number(matches[0]) : 0;
              }
              
              // ASIN情况处理
              let asinStr = null;
              if (asin && typeof asin === 'string') {
                asinStr = asin.trim();
                if (asinStr.length < 2) {
                  asinStr = null; // 无效ASIN视为null
                }
              }
              
              // 添加到映射
              if (asinStr) {
                // 如果同一个ASIN已存在，则累加值
                if (inventoryMap[asinStr] !== undefined) {
                  inventoryMap[asinStr] += inventory;
                  console.log(`ASIN ${asinStr} 在喜悦库存中出现多次，累加值为 ${inventoryMap[asinStr]}`);
                } else {
                  inventoryMap[asinStr] = inventory;
                  nameMap[asinStr] = nameStr;
                  foundCount++;
                }
              } else if (nameStr) {
                // 如果没有ASIN但有名称，使用自定义键
                const key = `NO_ASIN_${col}`;
                inventoryMap[key] = inventory;
                nameMap[key] = nameStr;
                // 这种情况也要记录为未匹配
                console.log(`列 ${col} 没有ASIN但有名称: ${nameStr}, 库存: ${inventory}`);
                foundCount++;
              }
            }
          }
          
          console.log(`成功从喜悦库存文件中找到${foundCount}个有效的产品库存数据`);
          console.log("ASIN样例:", Object.keys(inventoryMap).slice(0, 5));
          console.log("库存样例:", Object.values(inventoryMap).slice(0, 5));
          console.log("名称样例:", Object.values(nameMap).slice(0, 5));
          
          // 将数据存储到处理后的数据对象中
          processedData.value.inventoryMap = inventoryMap;
          processedData.value.nameMap = nameMap; // 保存名称映射
          resolve();
        } catch (error) {
          console.error("处理喜悦库存文件时出错:", error);
          reject(error);
        }
      };
      
      reader.onerror = function() {
        reject(new Error('读取文件时出错'));
      };
      
      reader.readAsArrayBuffer(file);
    } catch (error) {
      console.error("处理喜悦库存文件异常:", error);
      reject(error);
    }
  });
};

// 处理FBA库存文件
const processFBAInventory = async (file) => {
  return new Promise((resolve, reject) => {
    try {
      const reader = new FileReader();
      
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          
          // 查找表头行和关键列索引
          let headerRowIndex = -1;
          let asinColIndex = -1;
          let totalInventoryColIndex = -1;
          let inboundColIndex = -1;
          let unsellableColIndex = -1;
          let reservedColIndex = -1;
          
          // 查找表头行，通常是第一行或第二行
          for (let i = 0; i < Math.min(5, jsonData.length); i++) {
            if (jsonData[i] && jsonData[i].length > 5) {
              for (let j = 0; j < jsonData[i].length; j++) {
                const cellValue = String(jsonData[i][j] || '');
                
                // 找到ASIN列
                if (cellValue.includes("ASIN")) {
                  headerRowIndex = i;
                  asinColIndex = j;
                  break;
                }
              }
              if (headerRowIndex >= 0) break;
            }
          }
          
          if (headerRowIndex < 0) {
            reject(new Error('无法在FBA库存文件中找到表头行'));
            return;
          }
          
          // 查找必要的列
          const headerRow = jsonData[headerRowIndex];
          for (let i = 0; i < headerRow.length; i++) {
            const cellValue = String(headerRow[i] || '');
            
            if (cellValue.includes("总库存") || cellValue.includes("可售库存")) {
              totalInventoryColIndex = i;
              console.log(`找到总库存列: ${i+1}列`);
            } else if (cellValue.includes("入库处理中")) {
              inboundColIndex = i;
              console.log(`找到入库处理中列: ${i+1}列`);
            } else if (cellValue.includes("不可售总数")) {
              unsellableColIndex = i;
              console.log(`找到不可售列: ${i+1}列`);
            } else if (cellValue.includes("预留订单") || cellValue.includes("存货储备")) {
              reservedColIndex = i;
              console.log(`找到预留订单列: ${i+1}列`);
            }
          }
          
          // 确保找到了所有必要的列
          if (asinColIndex < 0 || totalInventoryColIndex < 0) {
            console.warn("未找到所有必要列，使用默认列索引");
            // 使用默认值（根据截图推测）
            if (asinColIndex < 0) asinColIndex = 2; // 假设ASIN在第3列
            if (totalInventoryColIndex < 0) totalInventoryColIndex = 11; // 假设总库存在第12列
            if (inboundColIndex < 0) inboundColIndex = 12; // 假设入库处理中在第13列
            if (unsellableColIndex < 0) unsellableColIndex = 14; // 假设不可售在第15列
            if (reservedColIndex < 0) reservedColIndex = 15; // 假设预留订单在第16列
          }
          
          // 创建ASIN到FBA库存计算值的映射
          const fbaInventoryMap = {};
          let foundCount = 0;
          
          // 从表头行的下一行开始处理
          for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length <= asinColIndex) continue;
            
            const asin = String(row[asinColIndex] || '').trim();
            if (!asin || asin.length < 5) continue; // 跳过无效ASIN
            
            // 获取各个值，默认为0
            const totalInventory = Number(row[totalInventoryColIndex] || 0);
            const inbound = Number(row[inboundColIndex] || 0);
            const unsellable = Number(row[unsellableColIndex] || 0);
            const reserved = Number(row[reservedColIndex] || 0);
            
            // 计算FBA库存值: 总库存 + 入库处理中 - 不可售总数 - 预留订单
            const fbaValue = totalInventory + inbound - unsellable - reserved;
            
            // 如果同一个ASIN已存在，则累加值，而不是覆盖
            if (fbaInventoryMap[asin] !== undefined) {
              fbaInventoryMap[asin] += fbaValue;
              console.log(`ASIN ${asin} 在FBA库存中出现多次，累加值为 ${fbaInventoryMap[asin]}`);
            } else {
              fbaInventoryMap[asin] = fbaValue;
              foundCount++;
            }
          }
          
          console.log(`成功从FBA库存文件中找到${foundCount}个有效的ASIN和库存数据`);
          console.log("FBA ASIN样例:", Object.keys(fbaInventoryMap).slice(0, 5));
          console.log("FBA值样例:", Object.values(fbaInventoryMap).slice(0, 5));
          
          // 将数据存储到处理后的数据对象中
          processedData.value.fbaInventoryMap = fbaInventoryMap;
          resolve();
        } catch (error) {
          console.error("处理FBA库存文件时出错:", error);
          reject(error);
        }
      };
      
      reader.onerror = function() {
        reject(new Error('读取FBA库存文件时出错'));
      };
      
      reader.readAsArrayBuffer(file);
    } catch (error) {
      console.error("处理FBA库存文件异常:", error);
      reject(error);
    }
  });
};

// 处理7天产品分析文件
const processSevenDayAnalysis = async (file) => {
  return new Promise((resolve, reject) => {
    try {
      const reader = new FileReader();
      
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          
          // ASIN列通常在第2列
          const asinColIndex = 1; // 第2列ASIN
          
          // 找到表头行并动态查找平均销量列
          let headerRowIndex = -1;
          let avgSalesColIndex = -1;
          
          for (let i = 0; i < Math.min(10, jsonData.length); i++) {
            if (jsonData[i] && jsonData[i].length > asinColIndex) {
              const asinCellValue = String(jsonData[i][asinColIndex] || '');
              if (asinCellValue.includes("ASIN")) {
                headerRowIndex = i;
                // 找到表头行后，查找平均销量列
                for (let j = 0; j < jsonData[i].length; j++) {
                  const headerValue = String(jsonData[i][j] || '');
                  if (headerValue.includes("平均销量") || headerValue.includes("Average Sales")) {
                    avgSalesColIndex = j;
                    console.log(`找到平均销量列，索引: ${j+1}列`);
                    break;
                  }
                }
                console.log(`找到7天分析表头行: ${i+1}行`);
                break;
              }
            }
          }
          
          // 如果未找到平均销量列，使用默认的第8列(H列)
          if (avgSalesColIndex < 0) {
            avgSalesColIndex = 7; // 默认使用第8列
            console.log(`未找到平均销量列，使用默认第8列(H列)`);
          }
          
          console.log(`使用列索引: ASIN列=${asinColIndex+1}, 平均销量列=${avgSalesColIndex+1}`);
          
          // 如果找不到表头行，假设第一行是表头
          if (headerRowIndex < 0) {
            headerRowIndex = 0;
            console.log("未找到明确的表头行，假设第一行是表头");
          }
          
          // 创建ASIN到7天平均销量的映射
          const sevenDaySalesMap = {};
          let foundCount = 0;
          
          // 从表头行的下一行开始处理
          for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length <= asinColIndex) continue;
            
            const asin = String(row[asinColIndex] || '').trim();
            if (!asin || asin.length < 5) continue; // 跳过无效ASIN
            
            // 获取平均销量值，默认为0
            let avgSales = row[avgSalesColIndex];
            if (avgSales === undefined || avgSales === null || isNaN(Number(avgSales))) {
              avgSales = 0;
            } else {
              avgSales = Number(avgSales);
            }
            
            // 如果同一个ASIN已存在，取较大值
            if (sevenDaySalesMap[asin] !== undefined) {
              sevenDaySalesMap[asin] = Math.max(sevenDaySalesMap[asin], avgSales);
              console.log(`ASIN ${asin} 在7天产品分析中出现多次，取较大值 ${sevenDaySalesMap[asin]}`);
            } else {
              sevenDaySalesMap[asin] = avgSales;
              foundCount++;
            }
          }
          
          console.log(`成功从7天产品分析文件中找到${foundCount}个ASIN的平均销量数据`);
          console.log("7天ASIN样例:", Object.keys(sevenDaySalesMap).slice(0, 5));
          console.log("7天平均销量样例:", Object.values(sevenDaySalesMap).slice(0, 5));
          
          // 将数据存储到处理后的数据对象中
          processedData.value.sevenDaySalesMap = sevenDaySalesMap;
          resolve();
        } catch (error) {
          console.error("处理7天产品分析文件时出错:", error);
          reject(error);
        }
      };
      
      reader.onerror = function() {
        reject(new Error('读取7天产品分析文件时出错'));
      };
      
      reader.readAsArrayBuffer(file);
    } catch (error) {
      console.error("处理7天产品分析文件异常:", error);
      reject(error);
    }
  });
};

// 处理30天产品分析文件
const processThirtyDayAnalysis = async (file) => {
  return new Promise((resolve, reject) => {
    try {
      const reader = new FileReader();
      
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          
          // ASIN列通常在第2列
          const asinColIndex = 1; // 第2列ASIN
          
          // 找到表头行并动态查找平均销量列
          let headerRowIndex = -1;
          let avgSalesColIndex = -1;
          
          for (let i = 0; i < Math.min(10, jsonData.length); i++) {
            if (jsonData[i] && jsonData[i].length > asinColIndex) {
              const asinCellValue = String(jsonData[i][asinColIndex] || '');
              if (asinCellValue.includes("ASIN")) {
                headerRowIndex = i;
                // 找到表头行后，查找平均销量列
                for (let j = 0; j < jsonData[i].length; j++) {
                  const headerValue = String(jsonData[i][j] || '');
                  if (headerValue.includes("平均销量") || headerValue.includes("Average Sales")) {
                    avgSalesColIndex = j;
                    console.log(`找到平均销量列，索引: ${j+1}列`);
                    break;
                  }
                }
                console.log(`找到30天分析表头行: ${i+1}行`);
                break;
              }
            }
          }
          
          // 如果未找到平均销量列，使用默认的第8列(H列)
          if (avgSalesColIndex < 0) {
            avgSalesColIndex = 7; // 默认使用第8列
            console.log(`未找到平均销量列，使用默认第8列(H列)`);
          }
          
          console.log(`使用列索引: ASIN列=${asinColIndex+1}, 平均销量列=${avgSalesColIndex+1}`);
          
          // 如果找不到表头行，假设第一行是表头
          if (headerRowIndex < 0) {
            headerRowIndex = 0;
            console.log("未找到明确的表头行，假设第一行是表头");
          }
          
          // 创建ASIN到30天平均销量的映射
          const thirtyDaySalesMap = {};
          let foundCount = 0;
          
          // 从表头行的下一行开始处理
          for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length <= asinColIndex) continue;
            
            const asin = String(row[asinColIndex] || '').trim();
            if (!asin || asin.length < 5) continue; // 跳过无效ASIN
            
            // 获取平均销量值，默认为0
            let avgSales = row[avgSalesColIndex];
            if (avgSales === undefined || avgSales === null || isNaN(Number(avgSales))) {
              avgSales = 0;
            } else {
              avgSales = Number(avgSales);
            }
            
            // 如果同一个ASIN已存在，取较大值
            if (thirtyDaySalesMap[asin] !== undefined) {
              thirtyDaySalesMap[asin] = Math.max(thirtyDaySalesMap[asin], avgSales);
              console.log(`ASIN ${asin} 在30天产品分析中出现多次，取较大值 ${thirtyDaySalesMap[asin]}`);
            } else {
              thirtyDaySalesMap[asin] = avgSales;
              foundCount++;
            }
          }
          
          console.log(`成功从30天产品分析文件中找到${foundCount}个ASIN的平均销量数据`);
          console.log("30天ASIN样例:", Object.keys(thirtyDaySalesMap).slice(0, 5));
          console.log("30天平均销量样例:", Object.values(thirtyDaySalesMap).slice(0, 5));
          
          // 将数据存储到处理后的数据对象中
          processedData.value.thirtyDaySalesMap = thirtyDaySalesMap;
          resolve();
        } catch (error) {
          console.error("处理30天产品分析文件时出错:", error);
          reject(error);
        }
      };
      
      reader.onerror = function() {
        reject(new Error('读取30天产品分析文件时出错'));
      };
      
      reader.readAsArrayBuffer(file);
    } catch (error) {
      console.error("处理30天产品分析文件异常:", error);
      reject(error);
    }
  });
};

// 生成结果文件 - 使用ExcelJS库
const generateResultFile = async () => {
  try {
    if (!baseFileData.value || !processedData.value.inventoryMap) {
      throw new Error('缺少必要的数据');
    }

    // 重置统计和未匹配ASIN列表
    processedData.value.statistics.totalFactoryUpdated = 0;
    processedData.value.statistics.totalFBAUpdated = 0;
    processedData.value.statistics.totalSevenDayUpdated = 0;
    processedData.value.statistics.totalThirtyDayUpdated = 0;
    processedData.value.unmatchedAsins = [];

    // 生成当前日期时间字符串
    const now = new Date();
    const dateStr = now.getFullYear() + 
      ('0' + (now.getMonth() + 1)).slice(-2) + 
      ('0' + now.getDate()).slice(-2) + '_' +
      ('0' + now.getHours()).slice(-2) + 
      ('0' + now.getMinutes()).slice(-2);
    
    // 结果文件名
    const resultFileName = `产品库存及周转统计_${dateStr}.xlsx`;
    
    console.log("开始使用ExcelJS处理文件...");
    
    try {
      // 获取原始文件
      const response = await fetch('/产品库存及周转统计.xlsx');
      if (!response.ok) {
        throw new Error(`获取文件失败: ${response.status} ${response.statusText}`);
      }
      
      const arrayBuffer = await response.arrayBuffer();
      console.log("原始文件大小:", arrayBuffer.byteLength, "字节");
      
      // 使用ExcelJS加载工作簿
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);
      console.log("成功加载工作簿, 工作表数量:", workbook.worksheets.length);
      
      // 获取第一个工作表
      const worksheet = workbook.worksheets[0];
      console.log("工作表名称:", worksheet.name, "行数:", worksheet.rowCount);
      
      // 收集所有有效的ASIN行及其ASIN
      const validRowNumbers = [];
      const foundAsins = new Set();
      
      // 步骤1：识别所有有效ASIN行
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) { // 从第2行开始
          const asinCell = row.getCell(2); // 第2列是ASIN列
          if (asinCell && asinCell.value) {
            const asin = String(asinCell.value).trim();
            // 检查是否是有效的ASIN (B开头，至少10个字符)
            if (asin.match(/^B[\dA-Z]{9,}/i)) {
              validRowNumbers.push(rowNumber);
              foundAsins.add(asin.toUpperCase());
            }
          }
        }
      });
      
      console.log(`找到${validRowNumbers.length}个有效的ASIN行`);
      
      // 步骤2：只处理有效ASIN行的工厂库存和FBA库存
      console.log("正在更新工厂库存、FBA库存和销量数据...");
      let updatedFactoryCount = 0;
      let updatedFBACount = 0;
      let updatedSevenDayCount = 0;
      let updatedThirtyDayCount = 0;
      
      // 遍历所有有效ASIN行
      for (const rowNumber of validRowNumbers) {
        const row = worksheet.getRow(rowNumber);
        const asin = String(row.getCell(2).value).trim().toUpperCase(); // ASIN
        
        // 更新工厂库存（第4列）
        const inventoryCell = row.getCell(4);
        if (processedData.value.inventoryMap && processedData.value.inventoryMap[asin] !== undefined) {
          const factoryInventory = processedData.value.inventoryMap[asin];
          inventoryCell.value = factoryInventory;
          updatedFactoryCount++;
        } else {
          // 如果没有匹配到喜悦库存数据，设置为0
          inventoryCell.value = 0;
          updatedFactoryCount++;
          console.log(`未找到 ASIN ${asin} 的工厂库存数据，设置为0`);
        }
        
        // 更新FBA库存（第3列）
        const fbaCell = row.getCell(3);
        if (processedData.value.fbaInventoryMap && processedData.value.fbaInventoryMap[asin] !== undefined) {
          const fbaInventory = processedData.value.fbaInventoryMap[asin];
          fbaCell.value = fbaInventory;
          updatedFBACount++;
        } else {
          // 如果没有匹配到FBA库存数据，设置为0
          fbaCell.value = 0;
          updatedFBACount++;
          console.log(`未找到 ASIN ${asin} 的FBA库存数据，设置为0`);
        }
        
        // 更新7天平均销量（第6列）
        const sevenDayCell = row.getCell(6);
        if (processedData.value.sevenDaySalesMap && processedData.value.sevenDaySalesMap[asin] !== undefined) {
          const sevenDaySales = processedData.value.sevenDaySalesMap[asin];
          sevenDayCell.value = sevenDaySales;
          updatedSevenDayCount++;
        } else {
          // 如果没有匹配到7天平均销量数据，设置为0
          sevenDayCell.value = 0;
          updatedSevenDayCount++;
          console.log(`未找到 ASIN ${asin} 的7天平均销量数据，设置为0`);
        }
        
        // 更新30天平均销量（第7列）
        const thirtyDayCell = row.getCell(7);
        if (processedData.value.thirtyDaySalesMap && processedData.value.thirtyDaySalesMap[asin] !== undefined) {
          const thirtyDaySales = processedData.value.thirtyDaySalesMap[asin];
          thirtyDayCell.value = thirtyDaySales;
          updatedThirtyDayCount++;
        } else {
          // 如果没有匹配到30天平均销量数据，设置为0
          thirtyDayCell.value = 0;
          updatedThirtyDayCount++;
          console.log(`未找到 ASIN ${asin} 的30天平均销量数据，设置为0`);
        }
        
        // 计算并更新总库存（第5列）
        const totalInventoryCell = row.getCell(5);
        const fbaInventory = Number(fbaCell.value || 0);
        const factoryInventory = Number(inventoryCell.value || 0);
        
        // 计算总和并设置值
        const totalInventory = fbaInventory + factoryInventory;
        totalInventoryCell.value = totalInventory;
        

        // 计算I列（最终计算每天平均）= MAX(F列,G列)*1.375
        const sevenDayAvg = Number(sevenDayCell.value || 0);  // F列(6)
        const thirtyDayAvg = Number(thirtyDayCell.value || 0); // G列(7)
        const finalAvgCell = row.getCell(9);  // I列
        const finalAvg = parseFloat((Math.max(sevenDayAvg, thirtyDayAvg) * 1.375).toFixed(2));
        finalAvgCell.value = finalAvg;

        // 计算J列（FBA周转）= C列/I列
        const fbaRotationCell = row.getCell(10);  // J列
        const fbaRotation = finalAvg > 0 ? fbaInventory / finalAvg : 0;
        fbaRotationCell.value = fbaRotation;

        // 计算K列（总周转）= E列/I列
        const totalRotationCell = row.getCell(11);  // K列
        const totalRotation = finalAvg > 0 ? totalInventory / finalAvg : 0;
        totalRotationCell.value = totalRotation;


        // 计算Q列（FBA库存总成本）= M列*C列
        const fbaCostCell = row.getCell(17);  // Q列
        const fbaCost = fbaCell.value * row.getCell(13).value;
        fbaCostCell.value = fbaCost;

        // 计算R列（工厂库存总成本）= N列*C列
        const factoryCostCell = row.getCell(18);  // R列
        const factoryCost = inventoryCell.value * row.getCell(13).value;
        factoryCostCell.value = factoryCost;
        

        // 获取T列(装箱数量)的值
        const tValue = Number(row.getCell(20).value || 36);

        // 计算U列 - 90天补货
        if(row.getCell(21).value !== null) {
          const uValue = Math.ceil(Math.max(0, 90 - Math.max(totalRotation, 60)) * finalAvg / tValue) * tValue;
          // 对于0值，使用一个特殊字符串来确保它显示在Excel中
          row.getCell(21).value = uValue === 0 ? "0" : uValue;
        }

        // 计算V列 - 120天补货(主推款)
        if(row.getCell(22).value !== null) {
          const vValue = Math.ceil(Math.max(0, 120 - Math.max(totalRotation, 60)) * finalAvg / tValue) * tValue;
          row.getCell(22).value = vValue === 0 ? "0" : vValue;
        }

        // 计算W列 - 135天补货(特推款)
        if(row.getCell(23).value !== null) {
          const wValue = Math.ceil(Math.max(0, 135 - Math.max(totalRotation, 60)) * finalAvg / tValue) * tValue;
          row.getCell(23).value = wValue === 0 ? "0" : wValue;
        }

        // 计算X列 - 70天发货
        if(row.getCell(24).value !== null) {
          const xBaseValue = Math.ceil(Math.max(0, 70 - Math.max(fbaRotation, 30)) * finalAvg / tValue) * tValue;
          const xValue = Math.min(xBaseValue, factoryInventory);
          row.getCell(24).value = xValue === 0 ? "0" : xValue;
        }

        // 计算Y列 - 100天发货(主推款)
        if(row.getCell(25).value !== null) {
          const yBaseValue = Math.ceil(Math.max(0, 100 - Math.max(fbaRotation, 30)) * finalAvg / tValue) * tValue;
          const yValue = Math.min(yBaseValue, factoryInventory);
          row.getCell(25).value = yValue === 0 ? "0" : yValue;
        }

        // 计算Z列 - 115天发货(特推款)
        if(row.getCell(26).value !== null) {
          const zBaseValue = Math.ceil(Math.max(0, 115 - Math.max(fbaRotation, 30)) * finalAvg / tValue) * tValue;
          const zValue = Math.min(zBaseValue, factoryInventory);
          row.getCell(26).value = zValue === 0 ? "0" : zValue;
        }
      }
      
      // 记录统计结果
      processedData.value.statistics.totalFactoryUpdated = updatedFactoryCount;
      processedData.value.statistics.totalFBAUpdated = updatedFBACount;
      processedData.value.statistics.totalSevenDayUpdated = updatedSevenDayCount;
      processedData.value.statistics.totalThirtyDayUpdated = updatedThirtyDayCount;
      
      console.log(`总共更新了 ${updatedFactoryCount} 个ASIN的工厂库存`);
      console.log(`总共更新了 ${updatedFBACount} 个ASIN的FBA库存`);
      console.log(`总共更新了 ${updatedSevenDayCount} 个ASIN的7天平均销量`);
      console.log(`总共更新了 ${updatedThirtyDayCount} 个ASIN的30天平均销量`);
      
      // 步骤3：查找喜悦库存中有但产品库存表中没有的ASIN
      if (processedData.value.inventoryMap && processedData.value.nameMap) {
        for (const key in processedData.value.inventoryMap) {
          const inventory = processedData.value.inventoryMap[key];
          const name = processedData.value.nameMap[key] || '';
          
          // 只添加库存数量大于0的未匹配产品
          if (inventory > 0) {
            // 检查是否是无ASIN的产品（以NO_ASIN_开头）
            if (key.startsWith('NO_ASIN_')) {
              // 这是没有ASIN但有名称的产品
              processedData.value.unmatchedAsins.push({
                asin: '', // 显示为空字符串
                name: name,
                inventory: inventory
              });
              console.log(`无ASIN产品 (${name})在喜悦库存中存在但在产品库存表中未找到，库存值: ${inventory}`);
            }
            // 检查是否是有ASIN的产品但在产品库存表中未找到
            else if (!foundAsins.has(key.toUpperCase())) {
              processedData.value.unmatchedAsins.push({
                asin: key,
                name: name,
                inventory: inventory
              });
              console.log(`ASIN ${key} (${name})在喜悦库存中存在但在产品库存表中未找到，库存值: ${inventory}`);
            }
          }
        }
      }
      
      console.log(`找到${processedData.value.unmatchedAsins.length}个未匹配的产品`);
      
      // 将工作簿转换为二进制数据
      const buffer = await workbook.xlsx.writeBuffer();
      
      // 保存结果
      resultFileData.value = {
        workbook: workbook,
        buffer: buffer,
        fileName: resultFileName
      };
      
      console.log("ExcelJS处理完成，生成了带样式的结果文件");
    } catch (error) {
      console.error("ExcelJS处理文件时出错:", error);
      throw error;
    }
  } catch (error) {
    console.error("生成结果文件时出错:", error);
    ElMessage({
      message: `生成结果文件时出错: ${error.message}`,
      type: 'error'
    });
  }

  // 移除这部分代码，改为在点击"下一步"按钮时执行
  // try {
  //   console.log("开始生成补货模版和发货模版...");
  //   await Promise.all([
  //     generateReplenishmentTemplate(),
  //     generateShippingTemplate(),
  //     generateBackendShippingTemplate()
  //   ]);
  //   console.log("所有模版生成完成");
  // } catch (templateError) {
  //   console.error("生成模版文件时出错:", templateError);
  //   ElMessage({
  //     message: `生成模版文件时出错: ${templateError.message}`,
  //   type: 'warning'
  //   });
  // }
};

// 生成补货模版
const generateReplenishmentTemplate = async () => {
  try {
    // 获取模版文件
    const response = await fetch('/补货模版.xlsx');
    if (!response.ok) {
      throw new Error(`获取补货模版失败: ${response.status} ${response.statusText}`);
    }
    
    const arrayBuffer = await response.arrayBuffer();
    console.log("补货模版文件大小:", arrayBuffer.byteLength, "字节");
    
    // 使用ExcelJS加载工作簿
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    console.log("成功加载补货模版工作簿, 工作表数量:", workbook.worksheets.length);
    
    // 获取第一个工作表
    const worksheet = workbook.worksheets[0];
    console.log("补货模版工作表名称:", worksheet.name, "行数:", worksheet.rowCount);
    
    // 从产品库存及周转统计中获取需要的数据
    if (!resultFileData.value || !resultFileData.value.workbook) {
      throw new Error('没有处理后的产品库存数据');
    }
    
    const inventoryWorkbook = resultFileData.value.workbook;
    const inventoryWorksheet = inventoryWorkbook.worksheets[0];
    
    // 找到补货模版的起始行
    let templateStartRow = 2; // 默认从第2行开始填充数据
    
    // 创建一个映射存储补货数据
    const replenishmentData = [];
    
    // 遍历产品库存表的每一行
    inventoryWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 1) { // 跳过表头
        const asinCell = row.getCell(2); // ASIN列
        if (asinCell && asinCell.value) {
          const asin = String(asinCell.value).trim();
          // 检查是否是有效的ASIN
          if (asin.match(/^B[\dA-Z]{9,}/i)) {
            // 获取需要的数据
            const productName = String(row.getCell(1).value || '').trim(); // 产品名称
            const tValue = Number(row.getCell(20).value || 0); // T列装箱数量
            
            // 获取UVW列的补货数据 - 这些是用于补货模版的
            const uValue = Number(row.getCell(21).value || 0); // U列 - 90天补货
            const vValue = Number(row.getCell(22).value || 0); // V列 - 主推款补货
            const wValue = Number(row.getCell(23).value || 0); // W列 - 特推款补货
            
            // FNSKU在AE列 (31列，因为索引从1开始)
            const fnsku = String(row.getCell(31).value || '').trim(); // AE列(FNSKU)
            
            // 只添加需要补货的行
            // 补货值选择UVW中最大值
            const maxReplenishmentValue = Math.max(uValue, vValue, wValue);
            
            if (maxReplenishmentValue > 0) {
              // 计算所需箱数
              const boxCount = tValue > 0 ? Math.ceil(maxReplenishmentValue / tValue) : 0;
              
              replenishmentData.push({
                name: productName,
                fnsku: fnsku,
                asin: asin,
                boxQuantity: tValue, // 每箱数量
                replenishmentQuantity: maxReplenishmentValue, // 补货数量
                boxCount: boxCount // 总箱数
              });
            }
          }
        }
      }
    });
    
    console.log(`找到${replenishmentData.length}个需要补货的产品`);
    
    // 填充补货模版
    let currentRow = templateStartRow;
    
    replenishmentData.forEach(item => {
      const row = worksheet.getRow(currentRow);
      
      // 设置单元格值 - 根据补货模版的实际结构可能需要调整列索引
      row.getCell(1).value = item.name; // 名称
      row.getCell(2).value = item.fnsku; // FNSKU
      row.getCell(3).value = item.boxQuantity; // 每箱数量
      row.getCell(4).value = item.replenishmentQuantity; // 补货数(套)
      row.getCell(5).value = item.boxCount; // 共多少箱
      
      currentRow++;
    });
    
    // 生成当前日期时间字符串
    const now = new Date();
    const dateStr = now.getFullYear() + 
      ('0' + (now.getMonth() + 1)).slice(-2) + 
      ('0' + now.getDate()).slice(-2) + '_' +
      ('0' + now.getHours()).slice(-2) + 
      ('0' + now.getMinutes()).slice(-2);
    
    // 结果文件名
    const resultFileName = `补货模版_${dateStr}.xlsx`;
    
    // 将工作簿转换为二进制数据
    const buffer = await workbook.xlsx.writeBuffer();
    
    // 保存结果
    processedTemplates.value.replenishmentTemplate = {
      workbook: workbook,
      buffer: buffer,
      fileName: resultFileName
    };
    
    console.log("补货模版处理完成");
    return true;
  } catch (error) {
    console.error("生成补货模版时出错:", error);
    throw error;
  }
};

// 生成发货模版
const generateShippingTemplate = async () => {
  try {
    // 获取模版文件
    const response = await fetch('/发货模版.xlsx');
    if (!response.ok) {
      throw new Error(`获取发货模版失败: ${response.status} ${response.statusText}`);
    }
    
    const arrayBuffer = await response.arrayBuffer();
    console.log("发货模版文件大小:", arrayBuffer.byteLength, "字节");
    
    // 使用ExcelJS加载工作簿
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    console.log("成功加载发货模版工作簿, 工作表数量:", workbook.worksheets.length);
    
    // 获取第一个工作表
    const worksheet = workbook.worksheets[0];
    console.log("发货模版工作表名称:", worksheet.name, "行数:", worksheet.rowCount);
    
    // 从产品库存及周转统计中获取需要的数据
    if (!resultFileData.value || !resultFileData.value.workbook) {
      throw new Error('没有处理后的产品库存数据');
    }
    
    const inventoryWorkbook = resultFileData.value.workbook;
    const inventoryWorksheet = inventoryWorkbook.worksheets[0];
    
    // 找到发货模版的起始行
    let templateStartRow = 2; // 默认从第2行开始填充数据
    
    // 创建一个映射存储发货数据
    const shippingData = [];
    
    // 遍历产品库存表的每一行
    inventoryWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 1) { // 跳过表头
        const asinCell = row.getCell(2); // ASIN列
        if (asinCell && asinCell.value) {
          const asin = String(asinCell.value).trim();
          // 检查是否是有效的ASIN
          if (asin.match(/^B[\dA-Z]{9,}/i)) {
            // 获取需要的数据
            const productName = String(row.getCell(1).value || '').trim(); // 产品名称
            const tValue = Number(row.getCell(20).value || 0); // T列装箱数量
            
            // 获取XYZ列的发货数据 - 这些是用于发货模版的
            const xValue = Number(row.getCell(24).value || 0); // X列
            const yValue = Number(row.getCell(25).value || 0); // Y列
            const zValue = Number(row.getCell(26).value || 0); // Z列
            
            // FNSKU在AE列 (31列，因为索引从1开始)
            const fnsku = String(row.getCell(31).value || '').trim(); // AE列(FNSKU)
            
            // 只添加需要发货的行
            // 发货值选择XYZ中最大值
            const maxShippingValue = Math.max(xValue, yValue, zValue);
            
            if (maxShippingValue > 0) {
              // 计算所需箱数
              const boxCount = tValue > 0 ? Math.ceil(maxShippingValue / tValue) : 0;
              
              shippingData.push({
                name: productName,
                fnsku: fnsku,
                asin: asin,
                boxQuantity: tValue, // 每箱数量
                shippingQuantity: maxShippingValue, // 发货数量
                boxCount: boxCount // 总箱数
              });
            }
          }
        }
      }
    });
    
    console.log(`找到${shippingData.length}个需要发货的产品`);
    
    // 填充发货模版
    let currentRow = templateStartRow;
    
    shippingData.forEach(item => {
      const row = worksheet.getRow(currentRow);
      
      // 设置单元格值 - 根据发货模版的实际结构可能需要调整列索引
      row.getCell(1).value = item.name; // 名称
      row.getCell(2).value = item.fnsku; // FNSKU
      row.getCell(3).value = item.boxQuantity; // 每箱数量
      row.getCell(4).value = item.shippingQuantity; // 发货数(套)
      row.getCell(5).value = item.boxCount; // 共多少箱
      
      currentRow++;
    });
    
    // 生成当前日期时间字符串
    const now = new Date();
    const dateStr = now.getFullYear() + 
      ('0' + (now.getMonth() + 1)).slice(-2) + 
      ('0' + now.getDate()).slice(-2) + '_' +
      ('0' + now.getHours()).slice(-2) + 
      ('0' + now.getMinutes()).slice(-2);
    
    // 结果文件名
    const resultFileName = `发货模版_${dateStr}.xlsx`;
    
    // 将工作簿转换为二进制数据
    const buffer = await workbook.xlsx.writeBuffer();
    
    // 保存结果
    processedTemplates.value.shippingTemplate = {
      workbook: workbook,
      buffer: buffer,
      fileName: resultFileName
    };
    
    console.log("发货模版处理完成");
    return true;
  } catch (error) {
    console.error("生成发货模版时出错:", error);
    throw error;
  }
};

// 生成后台发货模版
const generateBackendShippingTemplate = async () => {
  try {
    // 获取模版文件
    const response = await fetch('/后台发货模版.xlsx');
    if (!response.ok) {
      throw new Error(`获取后台发货模版失败: ${response.status} ${response.statusText}`);
    }
    
    const arrayBuffer = await response.arrayBuffer();
    console.log("后台发货模版文件大小:", arrayBuffer.byteLength, "字节");
    
    // 使用ExcelJS加载工作簿
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    console.log("成功加载后台发货模版工作簿, 工作表数量:", workbook.worksheets.length);
    
    // 获取"Create workflow – template"工作表
    const worksheetName = "Create workflow – template";
    const worksheet = workbook.getWorksheet(worksheetName);
    if (!worksheet) {
      throw new Error(`找不到工作表"${worksheetName}"`);
    }
    console.log(`后台发货模版工作表名称: ${worksheet.name}, 行数: ${worksheet.rowCount}`);
    
    // 从产品库存及周转统计中获取需要的数据
    if (!resultFileData.value || !resultFileData.value.workbook) {
      throw new Error('没有处理后的产品库存数据');
    }
    
    const inventoryWorkbook = resultFileData.value.workbook;
    const inventoryWorksheet = inventoryWorkbook.worksheets[0];
    
    // 找到后台发货模版的起始行 - 基于模版结构
    let templateStartRow = 9; // 后台模版从第9行开始填充数据
    
    // 创建一个映射存储发货数据
    const backendShippingData = [];
    
    // 遍历产品库存表的每一行
    inventoryWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 1) { // 跳过表头
        const asinCell = row.getCell(2); // ASIN列
        if (asinCell && asinCell.value) {
          const asin = String(asinCell.value).trim();
          // 检查是否是有效的ASIN
          if (asin.match(/^B[\dA-Z]{9,}/i)) {
            // 获取需要的数据
            const productName = String(row.getCell(1).value || '').trim(); // 产品名称
            const tValue = Number(row.getCell(20).value || 0); // T列装箱数量
            const sku = String(row.getCell(30).value || '').trim(); // AD列(SKU)
            
            // 获取XYZ列的发货数据 - 这些是用于发货模版的
            const xValue = Number(row.getCell(24).value || 0); // X列
            const yValue = Number(row.getCell(25).value || 0); // Y列
            const zValue = Number(row.getCell(26).value || 0); // Z列
            
            // 读取箱子尺寸和重量（调整为实际列数）
            // 需要将cm转换为英寸（除以2.54）
            const boxLengthCm = Number(row.getCell(32).value || 0); // 箱长 - AF列
            const boxWidthCm = Number(row.getCell(33).value || 0);  // 箱宽 - AG列  
            const boxHeightCm = Number(row.getCell(34).value || 0); // 箱高 - AH列
            
            // 将尺寸从cm转换为英寸
            const boxLength = boxLengthCm / 2.54;
            const boxWidth = boxWidthCm / 2.54;
            const boxHeight = boxHeightCm / 2.54;
            
            // 根据箱长确定箱重（lb）
            // 如果箱长为60，填45；其他情况填33
            const boxWeight = boxLengthCm === 60 ? 45 : 33;
            
            // 只添加需要发货的行
            // 发货值选择XYZ中最大值
            const maxShippingValue = Math.max(xValue, yValue, zValue);
            
            if (maxShippingValue > 0) {
              // 计算所需箱数
              const boxCount = tValue > 0 ? Math.ceil(maxShippingValue / tValue) : 0;
              
              backendShippingData.push({
                sku: sku,
                name: productName,
                asin: asin,
                quantity: maxShippingValue, // 发货数量
                unitsPerBox: tValue, // 每箱数量
                boxCount: boxCount, // 总箱数
                boxLength: boxLength.toFixed(2), // 箱长(英寸)，保留2位小数
                boxWidth: boxWidth.toFixed(2), // 箱宽(英寸)，保留2位小数
                boxHeight: boxHeight.toFixed(2), // 箱高(英寸)，保留2位小数
                boxWeight: boxWeight // 箱重(磅)
              });
            }
          }
        }
      }
    });
    
    console.log(`找到${backendShippingData.length}个需要添加到后台发货模版的产品`);
    
    // 填充后台发货模版
    let currentRow = templateStartRow;
    
    backendShippingData.forEach(item => {
      const row = worksheet.getRow(currentRow);
      
      // 设置单元格值 - 根据后台发货模版的实际结构
      row.getCell(1).value = item.sku; // Merchant SKU - AD列SKU
      row.getCell(2).value = item.quantity; // Quantity - 发货数量
      // Prep owner 和 Labeling owner 留空
      row.getCell(5).value = item.unitsPerBox; // Units per box - T列装箱数量
      row.getCell(6).value = item.boxCount; // Number of boxes - 发货数量/装箱数量
      row.getCell(7).value = item.boxLength; // Box length (in) - AF列/2.54
      row.getCell(8).value = item.boxWidth; // Box width (in) - AG列/2.54
      row.getCell(9).value = item.boxHeight; // Box height (in) - AH列/2.54
      row.getCell(10).value = item.boxWeight; // Box weight (lb) - 箱长为60填45，其他填33
      
      currentRow++;
    });
    
    // 生成当前日期时间字符串
    const now = new Date();
    const dateStr = now.getFullYear() + 
      ('0' + (now.getMonth() + 1)).slice(-2) + 
      ('0' + now.getDate()).slice(-2) + '_' +
      ('0' + now.getHours()).slice(-2) + 
      ('0' + now.getMinutes()).slice(-2);
    
    // 结果文件名
    const resultFileName = `后台发货模版_${dateStr}.xlsx`;
    
    // 将工作簿转换为二进制数据
    const buffer = await workbook.xlsx.writeBuffer();
    
    // 保存结果
    processedTemplates.value.backendShippingTemplate = {
      workbook: workbook,
      buffer: buffer,
      fileName: resultFileName
    };
    
    console.log("后台发货模版处理完成");
    return true;
  } catch (error) {
    console.error("生成后台发货模版时出错:", error);
    throw error;
  }
};

// 下载补货模版
const downloadReplenishmentTemplate = () => {
  try {
    if (!processedTemplates.value.replenishmentTemplate || !processedTemplates.value.replenishmentTemplate.buffer) {
    ElMessage({
        message: '没有可下载的补货模版',
    type: 'warning'
      });
      return;
    }
    
    // 创建Blob对象
    const blob = new Blob([processedTemplates.value.replenishmentTemplate.buffer], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    
    // 使用FileSaver.js下载文件
    FileSaver.saveAs(blob, processedTemplates.value.replenishmentTemplate.fileName);
    
    ElMessage({
      message: '补货模版下载成功！',
      type: 'success'
    });
  } catch (error) {
    ElMessage({
      message: `下载补货模版时出错: ${error.message}`,
      type: 'error'
    });
  }
};

// 下载发货模版
const downloadShippingTemplate = () => {
  try {
    if (!processedTemplates.value.shippingTemplate || !processedTemplates.value.shippingTemplate.buffer) {
    ElMessage({
        message: '没有可下载的发货模版',
    type: 'warning'
      });
      return;
    }
    
    // 创建Blob对象
    const blob = new Blob([processedTemplates.value.shippingTemplate.buffer], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    
    // 使用FileSaver.js下载文件
    FileSaver.saveAs(blob, processedTemplates.value.shippingTemplate.fileName);
    
    ElMessage({
      message: '发货模版下载成功！',
      type: 'success'
    });
  } catch (error) {
    ElMessage({
      message: `下载发货模版时出错: ${error.message}`,
      type: 'error'
    });
  }
};

// 下载后台发货模版
const downloadBackendShippingTemplate = () => {
  try {
    if (!processedTemplates.value.backendShippingTemplate || !processedTemplates.value.backendShippingTemplate.buffer) {
    ElMessage({
        message: '没有可下载的后台发货模版',
        type: 'warning'
      });
      return;
    }
    
    // 创建Blob对象
    const blob = new Blob([processedTemplates.value.backendShippingTemplate.buffer], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    
    // 使用FileSaver.js下载文件
    FileSaver.saveAs(blob, processedTemplates.value.backendShippingTemplate.fileName);
    
    ElMessage({
      message: '后台发货模版下载成功！',
      type: 'success'
    });
  } catch (error) {
    ElMessage({
      message: `下载后台发货模版时出错: ${error.message}`,
      type: 'error'
    });
  }
};

// 下载生成的文件 - 使用ExcelJS生成的文件
const downloadResultFile = () => {
  try {
    if (!resultFileData.value || !resultFileData.value.workbook) {
      ElMessage({
        message: '没有可下载的文件',
        type: 'warning'
      });
      return;
    }

    // 备份原始工作簿
    const originalWorkbook = resultFileData.value.workbook;
    
    // 克隆工作簿以避免修改原始数据
    originalWorkbook.xlsx.writeBuffer().then(buffer => {
      // 从buffer重新加载一个新的工作簿用于导出
      const workbook = new ExcelJS.Workbook();
      workbook.xlsx.load(buffer).then(() => {
        const worksheet = workbook.worksheets[0];
        
        console.log("正在准备导出文件...");
        
        // 对每行进行处理
        for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
          const row = worksheet.getRow(rowNumber);
          
          // 检查是否是有效的ASIN行
          const asinCell = row.getCell(2); // ASIN列
          if (!asinCell || !asinCell.value) continue;
          
          const asin = String(asinCell.value).trim();
          if (!asin.match(/^B[\dA-Z]{9,}/i)) continue;
          
          // 处理U-Z列（列索引21-26）
          for (let colIndex = 21; colIndex <= 26; colIndex++) {
            const cell = row.getCell(colIndex);
            
            // 检查单元格原始内容
            if (cell.value === null || cell.value === undefined) {
              // 如果单元格确实为空，强制设置为空
              cell.value = null;
            } 
            else if (cell.value === 0 || (typeof cell.value === 'number' && Math.abs(cell.value) < 0.001)) {
              // 对于0值，设置为字符串"0"以确保显示
              cell.value = "0";
            }
          }
        }
        
        console.log("文件准备完成，开始生成下载...");
        
        // 生成下载文件
        workbook.xlsx.writeBuffer().then(finalBuffer => {
          // 创建Blob对象
          const blob = new Blob([finalBuffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
          });
          
          // 使用FileSaver.js下载文件
          FileSaver.saveAs(blob, resultFileData.value.fileName);
          
          ElMessage({
            message: '文件下载成功！',
            type: 'success'
          });
        }).catch(error => {
          console.error("生成最终Buffer时出错:", error);
          ElMessage({
            message: `生成文件时出错: ${error.message}`,
            type: 'error'
          });
        });
      }).catch(error => {
        console.error("从Buffer加载工作簿时出错:", error);
        ElMessage({
          message: `处理Excel文件时出错: ${error.message}`,
          type: 'error'
        });
      });
    }).catch(error => {
      console.error("生成中间Buffer时出错:", error);
      ElMessage({
        message: `生成中间文件时出错: ${error.message}`,
        type: 'error'
      });
    });
  } catch (error) {
    console.error("下载文件过程中出错:", error);
    ElMessage({
      message: `下载文件时出错: ${error.message}`,
      type: 'error'
    });
  }
};

// 清理Excel中的空值，确保U-Z列的0值设置为空
const cleanEmptyCells = (workbook) => {
  try {
    const worksheet = workbook.worksheets[0];
    
    // 从第2行开始（跳过表头）
    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      
      // 检查是否是有效的ASIN行
      const asinCell = row.getCell(2); // ASIN列
      if (!asinCell || !asinCell.value) continue;
      
      const asin = String(asinCell.value).trim();
      if (!asin.match(/^B[\dA-Z]{9,}/i)) continue;
      
      // 检查原始模板中相应单元格是否为空
      // 通过一个简单的方法来判断：检查单元格的具体属性
      for (let colIndex = 21; colIndex <= 26; colIndex++) {
        const cell = row.getCell(colIndex);
        
        // 检查单元格的实际值，而不是指定值为null的行为
        let hasValue = false;
        
        // 如果单元格之前是由公式计算的，ExcelJS会保留一些特殊属性
        // cell.formula 或 cell.result 会有值，这表示它是通过公式计算的
        // 否则，就是直接设置的值
        
        if (cell.formula || (cell.value !== null && cell.value !== undefined && cell.value !== '')) {
          // 单元格有值
          hasValue = true;
        }
        
        // 检查单元格的值类型和内容
        if (!hasValue) {
          // 如果单元格原本就是空的，将其明确设置为undefined或null
          // 这确保了Excel导出时不会填充该单元格
          cell.value = undefined;
          console.log(`清理了 ASIN ${asin} 的第${colIndex}列空值`);
        } else if (cell.value === 0 || cell.value === '0' || (typeof cell.value === 'number' && Math.abs(cell.value) < 0.001)) {
          // 如果单元格值为0或几乎为0，确认它是否应该显示
          // 在这种情况下，我们想保留它作为显式的0
          if (cell.value !== "0") {
            cell.value = "0"; // 使用字符串0确保显示
            console.log(`将 ASIN ${asin} 的第${colIndex}列的0值转换为字符串"0"`);
          }
        }
      }
    }
    
    console.log('已完成空值清理');
  } catch (error) {
    console.error('清理空值时出错:', error);
  }
};

// 获取发货模版预览数据
const getShippingPreviewData = () => {
  if (!processedTemplates.value.shippingTemplate || !processedTemplates.value.shippingTemplate.workbook) {
    return [];
  }
  
  try {
    const worksheet = processedTemplates.value.shippingTemplate.workbook.worksheets[0];
    const dataRows = [];
    
    // 从第2行开始读取数据（跳过表头）
    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      
      // 读取名称、FNSKU等数据
      const name = row.getCell(1).value;
      // 如果名称为空，表示已经到了数据的末尾
      if (!name || name === '') {
        continue;
      }
      
      const fnsku = row.getCell(2).value;
      const boxQuantity = row.getCell(3).value;
      const shippingQuantity = row.getCell(4).value;
      const boxCount = row.getCell(5).value;
      
      dataRows.push({
        name: name,
        fnsku: fnsku,
        boxQuantity: boxQuantity,
        shippingQuantity: shippingQuantity,
        boxCount: boxCount
      });
    }
    
    // 按照发货数量排序，降序排列
    return dataRows.sort((a, b) => b.shippingQuantity - a.shippingQuantity);
  } catch (error) {
    console.error('获取发货模版预览数据时出错:', error);
    return [];
  }
};

// 获取补货模版预览数据
const getReplenishmentPreviewData = () => {
  if (!processedTemplates.value.replenishmentTemplate || !processedTemplates.value.replenishmentTemplate.workbook) {
    return [];
  }
  
  try {
    const worksheet = processedTemplates.value.replenishmentTemplate.workbook.worksheets[0];
    const dataRows = [];
    
    // 从第2行开始读取数据（跳过表头）
    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      
      // 读取名称、FNSKU等数据
      const name = row.getCell(1).value;
      // 如果名称为空，表示已经到了数据的末尾
      if (!name || name === '') {
        continue;
      }
      
      const fnsku = row.getCell(2).value;
      const boxQuantity = row.getCell(3).value;
      const replenishmentQuantity = row.getCell(4).value;
      const boxCount = row.getCell(5).value;
      
      dataRows.push({
        name: name,
        fnsku: fnsku,
        boxQuantity: boxQuantity,
        replenishmentQuantity: replenishmentQuantity,
        boxCount: boxCount
      });
    }
    
    // 按照补货数量排序，降序排列
    return dataRows.sort((a, b) => b.replenishmentQuantity - a.replenishmentQuantity);
  } catch (error) {
    console.error('获取补货模版预览数据时出错:', error);
    return [];
  }
};

// 添加一个函数来获取详细的产品库存表数据（包含发货和补货数量）
const getInventoryTableData = () => {
  if (!resultFileData.value || !resultFileData.value.workbook) {
    return [];
  }
  
  try {
    const worksheet = resultFileData.value.workbook.worksheets[0];
    const dataRows = [];
    
    // 从第2行开始读取数据（跳过表头）
    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      
      // 检查是否是有效的ASIN行
      const asinCell = row.getCell(2); // ASIN列
      if (!asinCell || !asinCell.value) continue;
      
      const asin = String(asinCell.value).trim();
      if (!asin.match(/^B[\dA-Z]{9,}/i)) continue;
      
      // 读取产品数据
      const productName = row.getCell(1).value;
      const factoryInventory = Number(row.getCell(4).value || 0);
      const fbaInventory = Number(row.getCell(3).value || 0);
      const totalInventory = Number(row.getCell(5).value || 0);
      const sevenDayAvg = Number(row.getCell(6).value || 0);
      const thirtyDayAvg = Number(row.getCell(7).value || 0);
      const finalDailyAvg = Number(row.getCell(9).value || 0);
      const fbaRotation = Number(row.getCell(10).value || 0);
      const totalRotation = Number(row.getCell(11).value || 0);
      
      // 读取补货和发货数量 - 使用Number转换确保是数字
      // 如果单元格是空的或非数字，则使用0作为默认值
      const replenishment90Day = Number(row.getCell(21).value || 0);
      const replenishment120Day = Number(row.getCell(22).value || 0);
      const replenishment135Day = Number(row.getCell(23).value || 0);
      
      const shipping70Day = Number(row.getCell(24).value || 0);
      const shipping100Day = Number(row.getCell(25).value || 0);
      const shipping115Day = Number(row.getCell(26).value || 0);
      
      // 获取最大的补货和发货数量
      const maxReplenishment = Math.max(replenishment90Day, replenishment120Day, replenishment135Day) || 0;
      const maxShipping = Math.max(shipping70Day, shipping100Day, shipping115Day) || 0;
      
      dataRows.push({
        rowNumber, // 保存行号，用于更新数据
        productName,
        asin,
        factoryInventory,
        fbaInventory,
        totalInventory,
        sevenDayAvg,
        thirtyDayAvg,
        finalDailyAvg,
        fbaRotation,
        totalRotation,
        maxReplenishment,
        maxShipping
      });
    }
    
    // 按照30天平均销量从大到小排序
    dataRows.sort((a, b) => b.thirtyDayAvg - a.thirtyDayAvg);
    
    return dataRows;
  } catch (error) {
    console.error('获取产品库存表数据时出错:', error);
    return [];
  }
};

// 更新最终计算每天平均值，并重新计算周转等数据
const updateFinalDailyAverage = (item) => {
  if (!resultFileData.value || !resultFileData.value.workbook) {
    ElMessage.error('无法更新数据，结果文件不存在');
    return;
  }
  
  try {
    const worksheet = resultFileData.value.workbook.worksheets[0];
    const row = worksheet.getRow(item.rowNumber);
    
    // 解析输入的值，确保是数字并保留两位小数
    const parsedValue = parseFloat(parseFloat(item.finalDailyAvg).toFixed(2));
    item.finalDailyAvg = isNaN(parsedValue) ? 0 : parsedValue;
    
    // 更新I列 - 最终计算每天平均
    const finalAvgCell = row.getCell(9);
    finalAvgCell.value = item.finalDailyAvg;
    
    // 重新计算FBA周转 = FBA库存/最终每天平均
    const fbaRotation = item.finalDailyAvg > 0 ? item.fbaInventory / item.finalDailyAvg : 0;
    const fbaRotationCell = row.getCell(10);
    fbaRotationCell.value = fbaRotation;
    item.fbaRotation = fbaRotation;
    
    // 重新计算总周转 = 总库存/最终每天平均
    const totalRotation = item.finalDailyAvg > 0 ? item.totalInventory / item.finalDailyAvg : 0;
    const totalRotationCell = row.getCell(11);
    totalRotationCell.value = totalRotation;
    item.totalRotation = totalRotation;
    
    // 获取基本数据
    const fbaInventory = item.fbaInventory;
    const factoryInventory = item.factoryInventory;
    const tValue = Number(row.getCell(20).value || 36); // T列(装箱数量)
    
    // 重新计算U列 - 90天补货
    // 公式: =CEILING(MAX(0,90-MAX(K{row},60))*I{row}/T{row},1)*T{row}
    const uValue = Math.ceil(Math.max(0, 90 - Math.max(totalRotation, 60)) * item.finalDailyAvg / tValue) * tValue;
    // 只有当单元格原本有值时才更新
    if (row.getCell(21).value !== null) {
      row.getCell(21).value = uValue === 0 ? "0" : uValue;
    }
    
    // 重新计算V列 - 120天补货(主推款)
    // 公式: =CEILING(MAX(0,120-MAX(K{row},60))*I{row}/T{row},1)*T{row}
    const vValue = Math.ceil(Math.max(0, 120 - Math.max(totalRotation, 60)) * item.finalDailyAvg / tValue) * tValue;
    if (row.getCell(22).value !== null) {
      row.getCell(22).value = vValue === 0 ? "0" : vValue;
    }
    
    // 重新计算W列 - 135天补货(特推款)
    // 公式: =CEILING(MAX(0,135-MAX(K{row},60))*I{row}/T{row},1)*T{row}
    const wValue = Math.ceil(Math.max(0, 135 - Math.max(totalRotation, 60)) * item.finalDailyAvg / tValue) * tValue;
    if (row.getCell(23).value !== null) {
      row.getCell(23).value = wValue === 0 ? "0" : wValue;
    }
    
    // 重新计算X列 - 70天发货
    // 公式: =MIN(CEILING(MAX(0,70-MAX(J{row},30))*I{row}/T{row},1)*T{row},D{row})
    const xBaseValue = Math.ceil(Math.max(0, 70 - Math.max(fbaRotation, 30)) * item.finalDailyAvg / tValue) * tValue;
    const xValue = Math.min(xBaseValue, factoryInventory);
    if (row.getCell(24).value !== null) {
      row.getCell(24).value = xValue === 0 ? "0" : xValue;
    }
    
    // 重新计算Y列 - 100天发货(主推款)
    // 公式: =MIN(CEILING(MAX(0,100-MAX(J{row},30))*I{row}/T{row},1)*T{row},D{row})
    const yBaseValue = Math.ceil(Math.max(0, 100 - Math.max(fbaRotation, 30)) * item.finalDailyAvg / tValue) * tValue;
    const yValue = Math.min(yBaseValue, factoryInventory);
    if (row.getCell(25).value !== null) {
      row.getCell(25).value = yValue === 0 ? "0" : yValue;
    }
    
    // 重新计算Z列 - 115天发货(特推款)
    // 公式: =MIN(CEILING(MAX(0,115-MAX(J{row},30))*I{row}/T{row},1)*T{row},D{row})
    const zBaseValue = Math.ceil(Math.max(0, 115 - Math.max(fbaRotation, 30)) * item.finalDailyAvg / tValue) * tValue;
    const zValue = Math.min(zBaseValue, factoryInventory);
    if (row.getCell(26).value !== null) {
      row.getCell(26).value = zValue === 0 ? "0" : zValue;
    }
    
    // 更新结果文件的buffer
    resultFileData.value.workbook.xlsx.writeBuffer().then(buffer => {
      resultFileData.value.buffer = buffer;
    });
    
    ElMessage.success('数据更新成功，已重新计算发补货数量');
  } catch (error) {
    console.error('更新数据时出错:', error);
    ElMessage.error(`更新数据时出错: ${error.message}`);
  }
};

// 进入发补货计划生成结果页面
const goToFinalResults = () => {
  // 重新生成所有模板
  Promise.all([
    generateReplenishmentTemplate(),
    generateShippingTemplate(),
    generateBackendShippingTemplate()
  ]).then(() => {
    uploadStep.value = 3; // 进入最终结果页面
    
    // 确保计算货值统计
    setTimeout(() => {
      calculateValues();
      console.log('在goToFinalResults函数中调用calculateValues');
    }, 500);
    
    ElMessage.success('模板生成成功，进入结果页面');
  }).catch(error => {
    ElMessage.error(`生成模板时出错: ${error.message}`);
  });
};

// 添加一个格式化数字的辅助函数
const formatNumber = (num) => {
  if (num === null || num === undefined) return 0;
  return parseFloat(num).toFixed(2);
};

// 删除发货模版中的项目
const deleteShippingItem = (row, index) => {
  if (!processedTemplates.value.shippingTemplate || !processedTemplates.value.shippingTemplate.workbook ||
      !processedTemplates.value.backendShippingTemplate || !processedTemplates.value.backendShippingTemplate.workbook) {
    ElMessage.error('无法删除，模板数据不完整');
    return;
  }

  try {
    // 1. 从发货模版中删除
    const shippingWorksheet = processedTemplates.value.shippingTemplate.workbook.worksheets[0];
    
    // 找到匹配的行并删除
    let shippingRowToDelete = -1;
    for (let rowNumber = 2; rowNumber <= shippingWorksheet.rowCount; rowNumber++) {
      const currentRow = shippingWorksheet.getRow(rowNumber);
      
      // 读取当前行的数据
      const name = currentRow.getCell(1).value;
      if (!name) continue;
      
      const fnsku = currentRow.getCell(2).value;
      
      // 通过比较名称和FNSKU确定是否是要删除的行
      if (String(name) === String(row.name) && String(fnsku) === String(row.fnsku)) {
        shippingRowToDelete = rowNumber;
        break;
      }
    }
    
    // 如果找到了匹配的行，删除它并移动后面的行
    if (shippingRowToDelete > 0) {
      shippingWorksheet.spliceRows(shippingRowToDelete, 1);
    }
    
    // 2. 从后台发货模版中删除
    const backendWorksheet = processedTemplates.value.backendShippingTemplate.workbook.getWorksheet("Create workflow – template");
    if (!backendWorksheet) {
      throw new Error('找不到后台发货模版的工作表');
    }
    
    // 在后台发货模版中查找匹配的行 (通过SKU和数量比较)
    let backendRowToDelete = -1;
    // 后台模版从第9行开始
    for (let rowNumber = 9; rowNumber <= backendWorksheet.rowCount; rowNumber++) {
      const currentRow = backendWorksheet.getRow(rowNumber);
      const merchantSku = currentRow.getCell(1).value;
      const quantity = currentRow.getCell(2).value;
      const unitsPerBox = currentRow.getCell(5).value;
      
      // 如果SKU为空，跳过该行
      if (!merchantSku) continue;
      
      // 比较数量和每箱数量确定是否是要删除的行
      if (Number(quantity) === Number(row.shippingQuantity) && Number(unitsPerBox) === Number(row.boxQuantity)) {
        backendRowToDelete = rowNumber;
        break;
      }
    }
    
    // 如果找到了匹配的行，删除它
    if (backendRowToDelete > 0) {
      backendWorksheet.spliceRows(backendRowToDelete, 1);
    }
    
    // 3. 更新两个模版的buffer
    Promise.all([
      processedTemplates.value.shippingTemplate.workbook.xlsx.writeBuffer(),
      processedTemplates.value.backendShippingTemplate.workbook.xlsx.writeBuffer()
    ]).then(([shippingBuffer, backendBuffer]) => {
      processedTemplates.value.shippingTemplate.buffer = shippingBuffer;
      processedTemplates.value.backendShippingTemplate.buffer = backendBuffer;
      
      ElMessage.success('成功删除发货项!');
    });
  } catch (error) {
    console.error('删除发货项时出错:', error);
    ElMessage.error(`删除发货项时出错: ${error.message}`);
  }
};

// 将发货项调整为5箱
const makeItFiveBoxes = (row, index) => {
  if (!processedTemplates.value.shippingTemplate || !processedTemplates.value.shippingTemplate.workbook ||
      !processedTemplates.value.backendShippingTemplate || !processedTemplates.value.backendShippingTemplate.workbook) {
    ElMessage.error('无法调整，模板数据不完整');
    return;
  }

  try {
    // 1. 更新发货模版中的数据
    const shippingWorksheet = processedTemplates.value.shippingTemplate.workbook.worksheets[0];
    
    // 找到匹配的行
    let shippingRowToUpdate = -1;
    for (let rowNumber = 2; rowNumber <= shippingWorksheet.rowCount; rowNumber++) {
      const currentRow = shippingWorksheet.getRow(rowNumber);
      
      // 读取当前行的数据
      const name = currentRow.getCell(1).value;
      if (!name) continue;
      
      const fnsku = currentRow.getCell(2).value;
      
      // 通过比较名称和FNSKU确定是否是要更新的行
      if (String(name) === String(row.name) && String(fnsku) === String(row.fnsku)) {
        shippingRowToUpdate = rowNumber;
        break;
      }
    }
    
    // 如果找到了匹配的行，更新箱数和发货数量
    if (shippingRowToUpdate > 0) {
      const currentRow = shippingWorksheet.getRow(shippingRowToUpdate);
      const boxQuantity = Number(currentRow.getCell(3).value || 0);
      
      // 更新发货数量为 boxQuantity * 5
      const newShippingQuantity = boxQuantity * 5;
      currentRow.getCell(4).value = newShippingQuantity;
      
      // 更新总箱数为5
      currentRow.getCell(5).value = 5;
    }
    
    // 2. 从后台发货模版中更新
    const backendWorksheet = processedTemplates.value.backendShippingTemplate.workbook.getWorksheet("Create workflow – template");
    if (!backendWorksheet) {
      throw new Error('找不到后台发货模版的工作表');
    }
    
    // 在后台发货模版中查找匹配的行 (通过SKU和数量比较)
    let backendRowToUpdate = -1;
    // 后台模版从第9行开始
    for (let rowNumber = 9; rowNumber <= backendWorksheet.rowCount; rowNumber++) {
      const currentRow = backendWorksheet.getRow(rowNumber);
      const merchantSku = currentRow.getCell(1).value;
      const quantity = currentRow.getCell(2).value;
      const unitsPerBox = currentRow.getCell(5).value;
      
      // 如果SKU为空，跳过该行
      if (!merchantSku) continue;
      
      // 比较数量和每箱数量确定是否是要更新的行
      if (Number(quantity) === Number(row.shippingQuantity) && Number(unitsPerBox) === Number(row.boxQuantity)) {
        backendRowToUpdate = rowNumber;
        break;
      }
    }
    
    // 如果找到了匹配的行，更新它
    if (backendRowToUpdate > 0) {
      const currentRow = backendWorksheet.getRow(backendRowToUpdate);
      const unitsPerBox = Number(currentRow.getCell(5).value || 0);
      
      // 更新发货数量为 unitsPerBox * 5
      const newShippingQuantity = unitsPerBox * 5;
      currentRow.getCell(2).value = newShippingQuantity;
      
      // 计算更新后的箱数 (列6是箱数)
      currentRow.getCell(6).value = 5;
    }
    
    // 3. 更新两个模版的buffer
    Promise.all([
      processedTemplates.value.shippingTemplate.workbook.xlsx.writeBuffer(),
      processedTemplates.value.backendShippingTemplate.workbook.xlsx.writeBuffer()
    ]).then(([shippingBuffer, backendBuffer]) => {
      processedTemplates.value.shippingTemplate.buffer = shippingBuffer;
      processedTemplates.value.backendShippingTemplate.buffer = backendBuffer;
      
      // 更新UI上的数据
      const updatedItems = getShippingPreviewData();
      processedData.value.shippingPreviewData = updatedItems;
      
      ElMessage.success('成功调整为5箱!');
    });
  } catch (error) {
    console.error('调整箱数时出错:', error);
    ElMessage.error(`调整箱数时出错: ${error.message}`);
  }
};

// 计算货值统计数据
const calculateValues = () => {
  console.log('开始执行calculateValues函数');
  
  if (!resultFileData.value || !resultFileData.value.workbook) {
    console.log('没有准备好的统计表数据，无法计算货值');
    return;
  }
  
  console.log('找到统计表数据，开始计算货值');
  
  try {
    // 从"统计表-下载产品库存及周转统计"中获取数据
    const workbook = resultFileData.value.workbook;
    const worksheet = workbook.worksheets[0];
    
    // 用于存储总和的变量
    let totalFbaCost = 0;
    let totalFactoryCost = 0;
    
    // 遍历表格，计算Q列(FBA库存总成本)和R列(工厂总成本)的总和
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) { // 跳过表头
        const fbaCost = Number(row.getCell(17).value || 0); // Q列(17)
        const factoryCost = Number(row.getCell(18).value || 0); // R列(18)
        
        totalFbaCost += fbaCost;
        totalFactoryCost += factoryCost;
      }
    });
    
    console.log(`FBA库存总成本(Q列)总和: ${totalFbaCost}`);
    console.log(`工厂总成本(R列)总和: ${totalFactoryCost}`);
    
    // 更新货值统计
    inventoryStats.value.amazonInventoryValue = totalFbaCost;
    inventoryStats.value.factoryInventoryValue = totalFactoryCost;
    
    // 更新统计信息文字
    updateStatsText();
    
  } catch (error) {
    console.error('计算货值时出错:', error);
  }
  
  console.log('calculateValues 执行完毕');
};

// 更新统计信息文字
const updateStatsText = () => {
  const now = new Date();
  const month = now.getMonth() + 1;
  const day = now.getDate();
  
  // 计算亚马逊货值（已经是人民币）
  const amazonValueRMB = inventoryStats.value.amazonInventoryValue;
  // 计算工厂实际货值（减去未付款）
  const factoryActualValue = inventoryStats.value.factoryInventoryValue - inventoryStats.value.factoryUnpaid;
  // 计算亚马逊预收款（人民币）
  const amazonAdvanceRMB = inventoryStats.value.amazonAdvance * inventoryStats.value.exchangeRate;
  // 计算银信致汇预收款（人民币）
  const yinxinAdvanceRMB = inventoryStats.value.yinxinAdvance * inventoryStats.value.exchangeRate;
  
  // 格式化为"xx万"形式的函数
  const formatToWan = (value) => {
    const wan = value / 10000;
    return wan.toFixed(2) + '万';
  };
  
  inventoryStats.value.statsText = `${month}月${day}日统计数据：亚马逊货值${formatToWan(amazonValueRMB)}，工厂货值${formatToWan(factoryActualValue)}，亚马逊预收款${formatToWan(amazonAdvanceRMB)}，银信致汇预收款${formatToWan(yinxinAdvanceRMB)}`;
};

// 创建响应式的统计数据
const inventoryStats = ref({
  amazonInventoryValue: 0, // 美元
  factoryInventoryValue: 0, // 人民币
  exchangeRate: 7.2, // 美元汇率
  factoryUnpaid: 0, // 工厂未付款（人民币）
  amazonAdvance: 0, // 亚马逊预收款（美元）
  yinxinAdvance: 0, // 银信致汇预收款（美元）
  statsText: '正在计算货值统计...' // 统计信息文字，设置默认值
});

// 监听用户输入变化，更新统计信息
watch(
  () => [
    inventoryStats.value.exchangeRate,
    inventoryStats.value.factoryUnpaid,
    inventoryStats.value.amazonAdvance,
    inventoryStats.value.yinxinAdvance
  ],
  () => {
    updateStatsText();
  },
  { deep: true }
);

// 当数据处理完毕后，计算货值
watch(
  () => resultFileData.value,
  (newValue) => {
    if (newValue) {
      calculateValues();
    }
  },
  { deep: true }
);

// 当上传步骤改变为第2步时，重新计算货值
watch(
  () => uploadStep.value,
  (newValue) => {
    if (newValue === 2) {
      calculateValues();
    }
  }
);

// 确保在页面加载时立即显示统计信息
onMounted(() => {
  updateStatsText();
});

// 当上传步骤改变时，在合适的时机计算货值
watch(
  () => uploadStep.value,
  (newValue) => {
    console.log(`uploadStep changed to ${newValue}`);
    if (newValue >= 2) {
      console.log('准备计算货值统计');
      // 使用setTimeout确保DOM已更新且数据已加载
      setTimeout(() => {
        calculateValues();
        console.log('在watch uploadStep中调用calculateValues');
      }, 500);
    }
  }
);
</script>

<template>
  <div class="app-container">
    <!-- 头部 -->
    <header class="app-header">
      <div class="logo-container">
        <div class="logo-icon">
          <el-icon><Box /></el-icon>
        </div>
        <h1>喜悦发补货计划</h1>
      </div>
      <div class="header-actions">
        <el-button type="text" @click="resetUploads" :disabled="fileList.length === 0">
          重置
        </el-button>
      </div>
    </header>
    
    <!-- 主要内容 -->
    <main class="app-content">
      <div v-if="uploadStep === 1" class="upload-section">
        <div class="intro-text">
          <h2>欢迎使用喜悦发补货计划系统</h2>
          <p>本系统用于每月统计销售情况、发补货计划，请按照要求上传以下文件：</p>
        </div>
        
        <!-- 文件下载提示 -->
        <div class="download-instructions">
          <h3>文件获取说明</h3>
          <div class="download-grid">
            <div class="download-item">
              <div class="download-icon"><el-icon><Document /></el-icon></div>
              <div class="download-content">
                <h4>喜悦库存文件</h4>
                <p>请从<a href="https://docs.qq.com/sheet/DZFhvT1lEc0pHRFBH" target="_blank">喜悦文档</a>下载，文件名需包含"喜悦"，保存为CSV格式</p>
              </div>
            </div>
            
            <div class="download-item">
              <div class="download-icon"><el-icon><Box /></el-icon></div>
              <div class="download-content">
                <h4>FBA库存文件</h4>
                <p>请从<a href="https://www.sellfox.com/amzup-web-main/web/inventoryManage/index.html" target="_blank">SellFox库存管理</a>下载，文件名需以"FBAInventory"开头</p>
              </div>
            </div>
            
            <div class="download-item">
              <div class="download-icon"><el-icon><DataAnalysis /></el-icon></div>
              <div class="download-content">
                <h4>7天产品分析</h4>
                <p>请从<a href="https://www.sellfox.com/amzup-web-main/web/data/productAnalysis/index.html" target="_blank">SellFox产品分析</a>下载，选择7-10天的日期范围</p>
              </div>
            </div>
            
            <div class="download-item">
              <div class="download-icon"><el-icon><DataLine /></el-icon></div>
              <div class="download-content">
                <h4>30天产品分析</h4>
                <p>请从<a href="https://www.sellfox.com/amzup-web-main/web/data/productAnalysis/index.html" target="_blank">SellFox产品分析</a>下载，选择30天以上的日期范围</p>
              </div>
            </div>
          </div>
        </div>
        
        <!-- 拖拽上传区域 -->
        <div 
          ref="dropZone" 
          class="drop-zone"
          @click="triggerFileSelect"
        >
          <div class="drop-icon">
            <el-icon><Upload /></el-icon>
          </div>
          <h3>拖拽文件到此处，或点击上传</h3>
          <p>支持同时上传多个文件，系统将自动识别文件类型</p>
          
          <div class="file-requirements">
            <div class="requirement-item">
              <el-icon><Document /></el-icon>
              <span>喜悦库存文件：名称包含"喜悦"的CSV文件</span>
            </div>
            <div class="requirement-item">
              <el-icon><Box /></el-icon>
              <span>FBA库存文件：以"FBAInventory"开头的Excel文件</span>
            </div>
            <div class="requirement-item">
              <el-icon><DataAnalysis /></el-icon>
              <span>7天产品分析：包含10天内日期范围的文件（如20250223-20250301）</span>
            </div>
            <div class="requirement-item">
              <el-icon><DataLine /></el-icon>
              <span>30天产品分析：包含30天以上日期范围的文件（如20250201-20250301）</span>
            </div>
          </div>
        </div>
        
        <!-- 上传状态卡片 -->
        <div class="upload-status-cards">
          <div class="status-card" :class="{ 'status-complete': fileReady.xiyue }">
            <div class="status-icon">
              <el-icon v-if="fileReady.xiyue"><Check /></el-icon>
              <el-icon v-else><Document /></el-icon>
            </div>
            <div class="status-text">
              <h4>喜悦库存文件</h4>
              <p v-if="fileReady.xiyue">已上传</p>
              <p v-else>待上传</p>
            </div>
          </div>
          
          <div class="status-card" :class="{ 'status-complete': fileReady.fba }">
            <div class="status-icon">
              <el-icon v-if="fileReady.fba"><Check /></el-icon>
              <el-icon v-else><Box /></el-icon>
            </div>
            <div class="status-text">
              <h4>FBA库存文件</h4>
              <p v-if="fileReady.fba">已上传</p>
              <p v-else>待上传</p>
            </div>
          </div>
          
          <div class="status-card" :class="{ 'status-complete': fileReady.sevenDay }">
            <div class="status-icon">
              <el-icon v-if="fileReady.sevenDay"><Check /></el-icon>
              <el-icon v-else><DataAnalysis /></el-icon>
            </div>
            <div class="status-text">
              <h4>7天产品分析</h4>
              <p v-if="fileReady.sevenDay">已上传</p>
              <p v-else>待上传</p>
            </div>
          </div>
          
          <div class="status-card" :class="{ 'status-complete': fileReady.thirtyDay }">
            <div class="status-icon">
              <el-icon v-if="fileReady.thirtyDay"><Check /></el-icon>
              <el-icon v-else><DataLine /></el-icon>
            </div>
            <div class="status-text">
              <h4>30天产品分析</h4>
              <p v-if="fileReady.thirtyDay">已上传</p>
              <p v-else>待上传</p>
            </div>
          </div>
        </div>
        
        <!-- 文件列表 -->
        <div class="uploaded-files-section" v-if="fileList.length > 0">
          <h3>已上传文件</h3>
          <el-table :data="fileList" style="width: 100%">
            <el-table-column prop="name" label="文件名" min-width="250"></el-table-column>
            <el-table-column prop="type" label="识别为" width="150">
              <template #default="scope">
                <span v-if="scope.row.type === 'xiyue'">喜悦库存</span>
                <span v-else-if="scope.row.type === 'fba'">FBA库存</span>
                <span v-else-if="scope.row.type === 'sevenDay'">7天产品分析</span>
                <span v-else-if="scope.row.type === 'thirtyDay'">30天产品分析</span>
              </template>
            </el-table-column>
            <el-table-column prop="size" label="文件大小" width="120"></el-table-column>
            <el-table-column prop="status" label="状态" width="100">
              <template #default="scope">
                <el-tag type="success" v-if="scope.row.status === 'uploaded'">已上传</el-tag>
                <el-tag type="warning" v-else>处理中</el-tag>
              </template>
            </el-table-column>
            <el-table-column label="操作" width="100">
              <template #default="scope">
                <el-button 
                  type="danger" 
                  size="small" 
                  circle
                  @click="removeFile(scope.row)"
                >
                  <el-icon><Delete /></el-icon>
                </el-button>
              </template>
            </el-table-column>
          </el-table>
        </div>
        
        <!-- 处理按钮 -->
        <div class="process-actions">
          <el-button 
            type="primary" 
            :disabled="!allFilesReady() || uploading" 
            @click="processFiles"
            :loading="uploading"
            size="large"
          >
            {{ uploading ? '处理中...' : '开始处理' }}
          </el-button>
        </div>
      </div>
      
      <div v-else-if="uploadStep === 2" class="results-section">
        <div class="results-header">
          <h2>产品库存及周转统计</h2>
          <p>已完成数据处理，您可以在这里查看和编辑"最终计算每天平均"数据</p>
        </div>
        
        <div class="inventory-table-section">
          <div class="table-instruction">
            <el-alert
              title="编辑说明"
              type="info"
              description="修改'最终计算每天平均'列的数值后，点击输入框外部或按回车键确认，系统将自动重新计算FBA周转、总周转以及发货和补货数量。"
              show-icon
              :closable="false"
              style="margin-bottom: 15px;"
            />
          </div>
          
          <el-table :data="getInventoryTableData()" style="width: 100%" height="450">
            <el-table-column prop="productName" label="产品名称" min-width="180" show-overflow-tooltip fixed="left"></el-table-column>
            <el-table-column prop="asin" label="ASIN" width="120"></el-table-column>
            <el-table-column prop="factoryInventory" label="工厂库存" width="100"></el-table-column>
            <el-table-column prop="fbaInventory" label="FBA库存" width="100"></el-table-column>
            <el-table-column prop="totalInventory" label="总库存" width="100"></el-table-column>
            <el-table-column prop="sevenDayAvg" label="7天平均" width="100"></el-table-column>
            <el-table-column prop="thirtyDayAvg" label="30天平均" width="100"></el-table-column>
            <el-table-column label="最终计算每天平均" width="180">
              <template #default="scope">
                <el-input 
                  v-model.number="scope.row.finalDailyAvg" 
                  @change="updateFinalDailyAverage(scope.row)"
                  type="number" 
                  step="0.01"
                  size="small"
                  class="final-avg-input"
                ></el-input>
              </template>
            </el-table-column>
            <el-table-column label="FBA周转" width="100">
              <template #default="scope">
                {{ formatNumber(scope.row.fbaRotation) }}
              </template>
            </el-table-column>
            <el-table-column label="总周转" width="100">
              <template #default="scope">
                {{ formatNumber(scope.row.totalRotation) }}
              </template>
            </el-table-column>
            <el-table-column label="补货数量" width="100">
              <template #default="scope">
                {{ scope.row.maxReplenishment }}
              </template>
            </el-table-column>
            <el-table-column label="发货数量" width="100">
              <template #default="scope">
                {{ scope.row.maxShipping }}
              </template>
            </el-table-column>
          </el-table>
          
          <div class="action-buttons" style="margin-top: 20px; text-align: center;">
            <el-button type="primary" size="large" @click="downloadResultFile">
              <el-icon><Download /></el-icon>
              下载产品库存及周转统计
            </el-button>
            <el-button type="success" size="large" @click="goToFinalResults">
              下一步：查看发补货计划
            </el-button>
          </div>
        </div>
      </div>
      
      <div v-else-if="uploadStep === 3" class="results-section">
        <div class="results-header">
          <h2>发补货计划生成结果</h2>
          <p>已完成文件处理，以下是重点数据预览</p>
        </div>
        
        <!-- 补货计划结果 -->
        <div class="results-placeholder">
          
          <!-- 未匹配ASIN列表 -->
          <div class="unmatched-asins-section" v-if="processedData.unmatchedAsins.length > 0">
            <h3 class="section-title">喜悦库存中存在但产品库存表中未找到的ASIN</h3>
            <el-table :data="processedData.unmatchedAsins" style="width: 100%; margin-top: 0.5rem;">
              <el-table-column prop="asin" label="ASIN" width="150"></el-table-column>
              <el-table-column prop="name" label="产品名称" min-width="300"></el-table-column>
              <el-table-column prop="inventory" label="库存数量" width="120"></el-table-column>
            </el-table>
          </div>
          
          <!-- 在未匹配产品信息下方添加模版预览表格 -->
          <div class="result-container">
            <!-- 模版预览部分 -->
            <div class="templates-preview-section">
              <!-- 发货预览表格 -->
              <div class="template-preview" v-if="processedTemplates.shippingTemplate">
                <h4 class="template-title">发货预览</h4>
                <div class="template-table-container full-width">
                  <table class="template-table">
                    <thead>
                      <tr>
                        <th width="25%">名称</th>
                        <th width="18%">FNSKU</th>
                        <th width="15%">每箱数量</th>
                        <th width="15%">发货数量</th>
                        <th width="17%">共多少箱</th>
                        <th width="10%">操作</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr v-for="(row, index) in getShippingPreviewData()" :key="index">
                        <td title="{{row.name}}">{{ row.name }}</td>
                        <td>{{ row.fnsku }}</td>
                        <td>{{ row.boxQuantity || 0 }}</td>
                        <td class="shipping-qty">{{ row.shippingQuantity || 0 }}</td>
                        <td>{{ row.boxCount || 0 }}</td>
                        <td>
                          <div class="action-buttons">
                            <el-button size="small" type="danger" @click="deleteShippingItem(row, index)">删除</el-button>
                            <el-button size="small" type="primary" @click="makeItFiveBoxes(row, index)">凑5箱</el-button>
                          </div>
                        </td>
                      </tr>
                      <tr v-if="getShippingPreviewData().length === 0">
                        <td colspan="6" class="no-data">没有需要发货的产品</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
          
          <div class="result-actions">
            <div class="action-group">
              <h4>统计表</h4>
              <el-button type="primary" size="large" @click="downloadResultFile">
                <el-icon><Download /></el-icon>
                下载产品库存及周转统计
              </el-button>
            </div>
            
            <div class="action-group">
              <h4>补货/发货模版</h4>
              <div class="template-buttons">
                <el-button type="success" size="large" @click="downloadReplenishmentTemplate">
                  <el-icon><Download /></el-icon>
                  下载补货模版
                </el-button>
                
                <el-button type="success" size="large" @click="downloadShippingTemplate">
                  <el-icon><Download /></el-icon>
                  下载发货模版
                </el-button>
                
                <el-button type="success" size="large" @click="downloadBackendShippingTemplate">
                  <el-icon><Download /></el-icon>
                  下载后台发货模版
                </el-button>
              </div>
            </div>
            
            <div class="action-group">
              <div class="buttons">
                <!-- 删除"下载结果处理文件"按钮，因为它与"下载产品库存及周转统计"功能重复 -->
                
                <!-- 货值统计 -->
                <div class="value-statistics-section">
                  <!-- 统计结果显示 -->
                  <div class="value-statistics-text">
                    {{ inventoryStats.statsText }}
                  </div>
                  
                  <!-- 简洁的输入框设计 -->
                  <div class="simple-inputs-row">
                    <div class="simple-input-group">
                      <label>美元汇率</label>
                      <el-input v-model="inventoryStats.exchangeRate" type="number" size="small" placeholder="美元汇率" />
                    </div>
                    <div class="simple-input-group">
                      <label>工厂未付款(¥)</label>
                      <el-input v-model="inventoryStats.factoryUnpaid" type="number" size="small" placeholder="工厂未付款" />
                    </div>
                    <div class="simple-input-group">
                      <label>亚马逊预收款($)</label>
                      <el-input v-model="inventoryStats.amazonAdvance" type="number" size="small" placeholder="亚马逊预收款" />
                    </div>
                    <div class="simple-input-group">
                      <label>银信致汇预收款($)</label>
                      <el-input v-model="inventoryStats.yinxinAdvance" type="number" size="small" placeholder="银信致汇预收款" />
                    </div>
                  </div>
                </div>
                
                <el-button @click="uploadStep = 1" size="large">返回上传页</el-button>
              </div>
            </div>
          </div>
        </div>
      </div>
    </main>
    
    <!-- 页脚 -->
    <footer class="app-footer">
      <p>© 2024 喜悦收纳整理 - 让生活更有条理</p>
    </footer>
  </div>
</template>

<style>
:root {
  --apple-blue: #0071e3;
  --apple-gray: #86868b;
  --apple-light-gray: #f5f5f7;
  --apple-off-white: #fbfbfd;
  --apple-dark: #1d1d1f;
  
  --home-primary: #f8d7da;
  --home-secondary: #d4edda;
  --home-accent: #ffeeba;
}

* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
  font-family: "SF Pro Display", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}

body {
  background-color: var(--apple-off-white);
  color: var(--apple-dark);
}

.app-container {
  display: flex;
  flex-direction: column;
  min-height: 100vh;
}

/* 头部样式 */
.app-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 1.5rem 2rem;
  background-color: rgba(255, 255, 255, 0.8);
  backdrop-filter: blur(20px);
  border-bottom: 1px solid rgba(0, 0, 0, 0.1);
  position: sticky;
  top: 0;
  z-index: 100;
}

.logo-container {
  display: flex;
  align-items: center;
}

.logo-icon {
  background-color: var(--apple-blue);
  color: white;
  width: 40px;
  height: 40px;
  border-radius: 10px;
  display: flex;
  align-items: center;
  justify-content: center;
  margin-right: 12px;
  font-size: 20px;
}

.app-header h1 {
  font-size: 1.5rem;
  font-weight: 600;
  color: var(--apple-dark);
}

/* 主内容样式 */
.app-content {
  flex: 1;
  padding: 2rem;
  max-width: 1200px;
  margin: 0 auto;
  width: 100%;
}

.intro-text {
  text-align: center;
  margin-bottom: 2rem;
}

.intro-text h2 {
  font-size: 2rem;
  margin-bottom: 1rem;
  font-weight: 600;
}

.intro-text p {
  color: var(--apple-gray);
  font-size: 1.1rem;
  max-width: 600px;
  margin: 0 auto;
}

/* 文件下载指导 */
.download-instructions {
  background-color: white;
  border-radius: 12px;
  padding: 1.5rem;
  margin-bottom: 2rem;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
}

.download-instructions h3 {
  margin-bottom: 1.2rem;
  font-weight: 600;
  text-align: center;
  color: var(--apple-dark);
}

.download-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(240px, 1fr));
  gap: 1.5rem;
}

.download-item {
  display: flex;
  align-items: flex-start;
}

.download-icon {
  width: 40px;
  height: 40px;
  border-radius: 8px;
  display: flex;
  align-items: center;
  justify-content: center;
  margin-right: 12px;
  flex-shrink: 0;
  font-size: 20px;
}

.download-item:nth-child(1) .download-icon {
  background-color: var(--home-primary);
  color: #721c24;
}

.download-item:nth-child(2) .download-icon {
  background-color: var(--home-secondary);
  color: #155724;
}

.download-item:nth-child(3) .download-icon,
.download-item:nth-child(4) .download-icon {
  background-color: var(--home-accent);
  color: #856404;
}

.download-content h4 {
  font-size: 1rem;
  font-weight: 600;
  margin-bottom: 0.5rem;
  color: var(--apple-dark);
}

.download-content p {
  font-size: 0.9rem;
  color: var(--apple-gray);
  line-height: 1.4;
}

.download-content a {
  color: var(--apple-blue);
  text-decoration: none;
  font-weight: 500;
}

.download-content a:hover {
  text-decoration: underline;
}

/* 拖拽区域样式 */
.drop-zone {
  border: 2px dashed #c0c4cc;
  border-radius: 12px;
  padding: 3rem 2rem;
  text-align: center;
  cursor: pointer;
  transition: all 0.3s ease;
  background-color: white;
  margin-bottom: 2rem;
}

.drop-zone:hover, .drop-zone.highlight {
  border-color: var(--apple-blue);
  background-color: rgba(0, 113, 227, 0.03);
}

.drop-icon {
  font-size: 48px;
  color: var(--apple-gray);
  margin-bottom: 1rem;
}

.drop-zone h3 {
  font-size: 1.5rem;
  margin-bottom: 0.5rem;
  font-weight: 600;
}

.drop-zone p {
  color: var(--apple-gray);
  margin-bottom: 2rem;
}

.file-requirements {
  max-width: 600px;
  margin: 0 auto;
  text-align: left;
}

.requirement-item {
  display: flex;
  align-items: center;
  margin-bottom: 0.75rem;
  color: var(--apple-gray);
}

.requirement-item .el-icon {
  margin-right: 0.75rem;
}

/* 上传状态卡片 */
.upload-status-cards {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(250px, 1fr));
  gap: 1rem;
  margin-bottom: 2rem;
}

.status-card {
  background-color: white;
  border-radius: 12px;
  padding: 1.25rem;
  display: flex;
  align-items: center;
  transition: all 0.3s ease;
  border: 1px solid #ebeef5;
}

.status-card.status-complete {
  background-color: rgba(103, 194, 58, 0.1);
  border-color: rgba(103, 194, 58, 0.2);
}

.status-icon {
  width: 40px;
  height: 40px;
  border-radius: 50%;
  background-color: #f0f2f5;
  display: flex;
  align-items: center;
  justify-content: center;
  margin-right: 1rem;
  font-size: 18px;
  color: var(--apple-gray);
}

.status-complete .status-icon {
  background-color: #67c23a;
  color: white;
}

.status-text h4 {
  font-size: 1rem;
  font-weight: 600;
  margin-bottom: 0.25rem;
}

.status-text p {
  color: var(--apple-gray);
  font-size: 0.9rem;
}

.status-complete .status-text p {
  color: #67c23a;
}

/* 已上传文件列表 */
.uploaded-files-section {
  background-color: white;
  border-radius: 12px;
  padding: 1.5rem;
  margin-bottom: 2rem;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
}

.uploaded-files-section h3 {
  margin-bottom: 1rem;
  font-weight: 600;
}

/* 处理按钮 */
.process-actions {
  display: flex;
  justify-content: center;
  margin-top: 2rem;
}

.process-actions button {
  min-width: 200px;
  height: 50px;
}

/* 结果部分 */
.results-section {
  text-align: center;
}

.results-header {
  margin-bottom: 2rem;
}

.results-placeholder {
  background-color: white;
  border-radius: 12px;
  padding: 3rem 2rem;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
  display: flex;
  flex-direction: column;
  align-items: center;
}

.placeholder-icon {
  width: 100px;
  height: 100px;
  background-color: var(--apple-light-gray);
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 48px;
  color: var(--apple-blue);
  margin-bottom: 1.5rem;
}

.results-placeholder h3 {
  font-size: 1.5rem;
  margin-bottom: 1rem;
}

.results-placeholder p {
  color: var(--apple-gray);
  margin-bottom: 2rem;
  max-width: 500px;
}

.result-details {
  display: flex;
  gap: 2rem;
  justify-content: center;
  margin-bottom: 2rem;
  flex-wrap: wrap;
}

.result-stat {
  text-align: center;
  background-color: var(--apple-light-gray);
  border-radius: 12px;
  padding: 1.5rem;
  min-width: 150px;
}

.result-stat h4 {
  font-size: 0.9rem;
  color: var(--apple-gray);
  margin-bottom: 0.5rem;
}

.stat-value {
  font-size: 2.5rem;
  font-weight: 600;
  color: var(--apple-blue);
  margin-bottom: 0.25rem;
}

.stat-label {
  font-size: 0.9rem;
  color: var(--apple-gray);
}

/* 未匹配ASIN列表样式 */
.unmatched-asins-section {
  width: 100%;
  margin: 1rem 0 2rem;
  background-color: rgba(255, 229, 229, 0.3);
  border-radius: 12px;
  padding: 1.5rem;
  border: 1px solid rgba(220, 53, 69, 0.2);
}

.unmatched-asins-section h3 {
  color: #721c24;
  font-size: 1.1rem;
  font-weight: 600;
  margin-bottom: 0.5rem;
}

.result-actions {
  display: flex;
  flex-direction: column;
  gap: 1.5rem;
  align-items: center;
  margin-top: 2rem;
  width: 100%;
}

.action-group {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 0.7rem;
  width: 100%;
  max-width: 800px;
}

.action-group h4 {
  font-size: 1.1rem;
  font-weight: 600;
  color: var(--apple-dark);
}

.template-buttons {
  display: flex;
  flex-wrap: wrap;
  gap: 1rem;
  justify-content: center;
}

@media (max-width: 768px) {
  .template-buttons {
    flex-direction: column;
  }
}

/* 页脚样式 */
.app-footer {
  background-color: var(--apple-light-gray);
  padding: 1.5rem;
  text-align: center;
  font-size: 0.9rem;
  color: var(--apple-gray);
  margin-top: 2rem;
}

/* 响应式调整 */
@media (max-width: 768px) {
  .upload-status-cards {
    grid-template-columns: 1fr;
  }
  
  .app-header {
    padding: 1rem;
  }
  
  .app-content {
    padding: 1rem;
  }
  
  .logo-icon {
    width: 36px;
    height: 36px;
  }
  
  .app-header h1 {
    font-size: 1.2rem;
  }
  
  .result-details {
    flex-direction: column;
    gap: 1rem;
    align-items: center;
  }
  
  .result-stat {
    width: 100%;
    max-width: 300px;
  }
}

/* 图标大小 */
.el-icon {
  font-size: inherit;
  vertical-align: middle;
}

.card-icon .el-icon {
  font-size: 24px;
}

.placeholder-icon .el-icon {
  font-size: 48px;
}

.template-stat {
  color: #67c23a; /* Success green color */
  font-size: 2rem;
  display: flex;
  justify-content: center;
  align-items: center;
  height: 60px;
}

.template-icon {
  font-size: 2.5rem;
  color: #67c23a;
  background-color: rgba(103, 194, 58, 0.1);
  border-radius: 50%;
  padding: 8px;
}

/* 表格容器样式 */
.template-preview {
  width: 100%;
  margin-top: 1.5rem;
}

.template-title {
  font-size: 1.1rem;
  font-weight: 600;
  color: #606266;
  margin-bottom: 0.5rem;
  text-align: center;
  background-color: #f0f9eb;
  padding: 8px;
  border-radius: 4px;
}

.template-table-container {
  border: 1px solid #ebeef5;
  border-radius: 4px;
}

.template-table-container.full-width {
  width: 100%;
  max-height: none;
  overflow: visible;
}

.template-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 14px;
}

.template-table th,
.template-table td {
  padding: 12px 8px;
  text-align: center;
  border-bottom: 1px solid #ebeef5;
}

.template-table th {
  background-color: #f5f7fa;
  color: #606266;
  font-weight: 600;
}

.template-table tr:hover {
  background-color: #f5f7fa;
}

.no-data {
  text-align: center;
  color: #909399;
  padding: 20px;
  font-style: italic;
}

.results-placeholder {
  margin-top: 1rem;
}

.unmatched-asins-section {
  margin-top: 1rem;
}

.section-title {
  font-size: 1.1rem;
  font-weight: 600;
  color: #606266;
  margin-bottom: 0.5rem;
}

.results-header {
  margin-bottom: 1rem;
}

.results-header h2 {
  font-size: 1.5rem;
  margin-bottom: 0.5rem;
}

.results-header p {
  color: #909399;
  margin: 0;
}

/* 发货预览表格样式 */
.shipping-qty {
  font-weight: 600;
  color: #409eff;
}

.template-table td {
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  max-width: 300px;
}

/* 结果页面样式 */
.results-section {
  padding: 2rem;
  min-height: calc(100vh - 160px);
}

.results-header {
  margin-bottom: 2rem;
  text-align: center;
}

.results-header h2 {
  font-size: 1.8rem;
  color: var(--apple-dark);
  margin-bottom: 0.5rem;
}

.results-header p {
  color: var(--apple-gray);
  font-size: 1rem;
}

.results-placeholder {
  background-color: white;
  border-radius: 8px;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
  padding: 2rem;
  margin-bottom: 2rem;
}

/* 产品库存表格页面样式 */
.inventory-table-section {
  background-color: white;
  border-radius: 8px;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
  padding: 1.5rem;
  margin-bottom: 2rem;
}

.inventory-table-section .el-table {
  --el-table-header-bg-color: #f5f7fa;
  --el-table-border-color: #ebeef5;
}

.inventory-table-section .el-table th {
  font-weight: 600;
  color: #606266;
}

.inventory-table-section .el-input {
  width: 100%;
}

.inventory-table-section .el-input__inner {
  text-align: center;
  font-weight: 500;
  color: #409EFF;
}

.action-buttons {
  display: flex;
  justify-content: center;
  gap: 1rem;
  margin-top: 1.5rem;
}

/* 修改现有样式 */
.inventory-table-section .final-avg-input .el-input__inner {
  text-align: center;
  font-weight: 500;
  color: #409EFF;
}

.inventory-table-section .el-table td {
  padding: 8px 0;
}

.inventory-table-section .el-table .cell {
  text-align: center;
}

/* 移除之前的首列左对齐设置 */
/* .inventory-table-section .el-table .el-table__cell:first-child .cell {
  text-align: left;
  padding-left: 12px;
} */

.action-buttons {
  display: flex;
  gap: 5px;
}

/* 简化货值统计样式 */
.value-statistics-text {
  margin: 20px 0;
  padding: 12px 15px;
  border-radius: 4px;
  background-color: #f0f8ff;
  color: #303133;
  font-size: 16px;
  line-height: 1.5;
  text-align: center;
  border-left: 4px solid #409eff;
  font-weight: 500;
}

/* 删除旧的货值统计样式 */
.value-statistics-section {
  margin: 20px 0;
  border-radius: 4px;
  background-color: #f0f8ff;
  border-left: 4px solid #409eff;
  padding: 10px 15px;
}

.value-statistics-text {
  color: #303133;
  font-size: 16px;
  line-height: 1.5;
  text-align: center;
  font-weight: 500;
  margin-bottom: 12px;
}

/* 简洁的输入框行样式 */
.simple-inputs-row {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  justify-content: center;
}

.simple-input-group {
  display: flex;
  flex-direction: column;
  min-width: 160px;
}

.simple-input-group label {
  font-size: 13px;
  color: #606266;
  margin-bottom: 4px;
}

/* 删除旧样式 */
.statistics-form {
  margin-top: 15px;
}

.form-fields {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 15px;
  margin-bottom: 20px;
}

.form-group {
  display: flex;
  flex-direction: column;
}

.form-group label {
  margin-bottom: 5px;
  font-size: 14px;
  color: #606266;
}

.statistics-result {
  background-color: #f0f8ff;
  padding: 15px;
  border-radius: 4px;
  border-left: 4px solid #409eff;
}

.result-text {
  margin: 0;
  font-size: 15px;
  line-height: 1.5;
  color: #303133;
  font-weight: 500;
}

/* 调整返回按钮位置 */
.back-button-container {
  margin-top: 20px;
  display: flex;
  justify-content: center;
}
</style>
