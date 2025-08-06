// 添加错误处理以捕获模块导入问题
let ExcelProcessor, DataVisualizer;

// 尝试导入模块
import('./excelProcessor.js')
  .then(module => {
    ExcelProcessor = module.ExcelProcessor;
    console.log('ExcelProcessor模块导入成功');
    checkModulesLoaded();
  })
  .catch(error => {
    console.error('ExcelProcessor模块导入失败:', error);
  });

import('./dataVisualizer.js')
  .then(module => {
    DataVisualizer = module.DataVisualizer;
    console.log('DataVisualizer模块导入成功');
    checkModulesLoaded();
  })
  .catch(error => {
    console.error('DataVisualizer模块导入失败:', error);
  });

// 检查模块是否都已加载
let modulesLoaded = false;
function checkModulesLoaded() {
  if (ExcelProcessor && DataVisualizer && !modulesLoaded) {
    modulesLoaded = true;
    console.log('所有模块加载完成，初始化应用');
    initializeApp();
  }
}

// DOM元素
let fileInput, filterSection, viewSection, dataSection, filterControls, 
    viewControls, paginationControls, dataContainer, dataInfo, debugSection, debugInfo;

// 全局变量
let originalData = [];
let currentData = [];
let headers = [];
let currentIndex = 0;
let currentView = 'single'; // 'single', 'json', 'table'

// 显示调试信息的函数
function showDebugInfo(message) {
  if (debugInfo) {
    const time = new Date().toLocaleTimeString();
    debugInfo.innerHTML += `<p>[${time}] ${message}</p>`;
    debugInfo.scrollTop = debugInfo.scrollHeight; // 滚动到底部
  }
  if (debugSection) {
    debugSection.classList.remove('hidden');
  }
  console.log(message);
}

// 显示加载状态的函数
function showLoadingStatus(message) {
  if (dataContainer) {
    dataContainer.innerHTML = `<p>⏳ ${message}</p>`;
  }
}

// 初始化应用
function initializeApp() {
  showDebugInfo('开始初始化应用');
  
  // 获取DOM元素
  fileInput = document.getElementById('file-input');
  filterSection = document.getElementById('filter-section');
  viewSection = document.getElementById('view-section');
  dataSection = document.getElementById('data-section');
  filterControls = document.getElementById('filter-controls');
  viewControls = document.getElementById('view-controls');
  paginationControls = document.getElementById('pagination-controls');
  dataContainer = document.getElementById('data-container');
  dataInfo = document.getElementById('data-info');
  debugSection = document.getElementById('debug-section');
  debugInfo = document.getElementById('debug-info');
  
  // 检查DOM元素是否存在
  if (!fileInput) {
    showDebugInfo('错误: 无法找到文件输入元素');
    return;
  }
  
  showDebugInfo('DOM元素获取完成');
  
  // 绑定事件监听器
  fileInput.addEventListener('change', handleFileSelect);
  
  showDebugInfo('应用初始化完成');
}

// 文件上传处理
async function handleFileSelect(event) {
  showDebugInfo('文件选择事件触发');
  
  const file = event.target.files[0];
  if (!file) {
    showDebugInfo('未选择文件');
    return;
  }
  
  showDebugInfo(`选择的文件: ${file.name}, 大小: ${file.size} bytes`);
  
  try {
    // 检查模块是否已加载
    if (!ExcelProcessor || !DataVisualizer) {
      showDebugInfo('错误: 必需的模块未加载');
      if (dataContainer) {
        dataContainer.innerHTML = '<p class="error">应用初始化未完成，请刷新页面后重试</p>';
      }
      return;
    }
    
    // 显示加载状态
    showLoadingStatus('正在读取Excel文件...');
    if (dataSection) {
      dataSection.classList.remove('hidden');
    }
    
    showDebugInfo('开始读取Excel文件');
    // 读取Excel文件
    const rawData = await ExcelProcessor.readFile(file);
    showDebugInfo(`原始数据读取完成，数据行数: ${rawData ? rawData.length : 0}`);
    
    // 处理数据
    showLoadingStatus('正在处理数据...');
    showDebugInfo('开始处理数据');
    const processedData = ExcelProcessor.processData(rawData);
    showDebugInfo(`数据处理完成，处理后的数据条数: ${processedData.data ? processedData.data.length : 0}`);
    
    originalData = processedData.data || [];
    currentData = [...originalData];
    headers = processedData.headers || [];
    
    showDebugInfo(`数据赋值完成，原始数据条数: ${originalData.length}`);
    
    // 重置索引
    currentIndex = 0;
    
    // 显示筛选控件
    if (filterControls && originalData && headers) {
      showLoadingStatus('正在创建筛选控件...');
      showDebugInfo('创建筛选控件');
      DataVisualizer.createFilterControls(
        headers, 
        originalData,
        filterControls, 
        handleFilterChange
      );
    }
    
    if (filterSection) {
      filterSection.classList.remove('hidden');
    }
    
    // 显示视图控制
    if (viewControls) {
      showLoadingStatus('正在创建视图控制...');
      showDebugInfo('创建视图控制');
      DataVisualizer.createViewControls(
        currentView,
        handleViewChange,
        viewControls
      );
    }
    
    if (viewSection) {
      viewSection.classList.remove('hidden');
    }
    
    // 显示数据
    showLoadingStatus('正在显示数据...');
    showDebugInfo('显示数据信息');
    displayDataInfo();
    showDebugInfo('显示当前数据');
    displayCurrentData();
    
    showDebugInfo('文件处理完成');
  } catch (error) {
    const errorMsg = `处理文件时出错: ${error.message}\n堆栈信息: ${error.stack}`;
    showDebugInfo(errorMsg);
    console.error('处理文件时出错:', error);
    if (dataContainer) {
      dataContainer.innerHTML = `<p class="error">处理文件时出错: ${error.message}</p>
                                 <p>请检查文件格式是否正确（支持 .xlsx 和 .xls 格式）</p>
                                 <p>详细错误信息已在调试区域显示</p>`;
    }
  }
}

// 筛选条件变化处理
function handleFilterChange(filters) {
  showDebugInfo(`筛选条件变化: ${JSON.stringify(filters)}`);
  
  try {
    // 检查模块是否已加载
    if (!ExcelProcessor) {
      showDebugInfo('错误: ExcelProcessor模块未加载');
      return;
    }
    
    // 显示加载状态
    showLoadingStatus('正在筛选数据...');
    
    // 应用筛选
    if (originalData) {
      const oldLength = currentData.length;
      currentData = ExcelProcessor.filterData(originalData, filters);
      showDebugInfo(`筛选完成，筛选前: ${oldLength} 条，筛选后: ${currentData.length} 条`);
    }
    
    // 重置索引
    currentIndex = 0;
    
    // 更新显示
    displayDataInfo();
    displayCurrentData();
    
    // 更新筛选控件显示
    if (filterControls && DataVisualizer) {
      DataVisualizer.updateFilterControls(filters, filterControls);
    }
  } catch (error) {
    const errorMsg = `处理筛选条件变化时出错: ${error.message}\n堆栈信息: ${error.stack}`;
    showDebugInfo(errorMsg);
    console.error('处理筛选条件变化时出错:', error);
  }
}

// 视图变化处理
function handleViewChange(view) {
  showDebugInfo(`视图变化: ${view}`);
  
  try {
    // 检查模块是否已加载
    if (!DataVisualizer) {
      showDebugInfo('错误: DataVisualizer模块未加载');
      return;
    }
    
    currentView = view;
    
    // 显示加载状态
    showLoadingStatus('正在切换视图...');
    
    // 更新视图控制按钮状态
    if (viewControls) {
      DataVisualizer.createViewControls(
        currentView,
        handleViewChange,
        viewControls
      );
    }
    
    // 重新显示当前数据
    displayCurrentData();
  } catch (error) {
    const errorMsg = `处理视图变化时出错: ${error.message}\n堆栈信息: ${error.stack}`;
    showDebugInfo(errorMsg);
    console.error('处理视图变化时出错:', error);
  }
}

// 显示当前数据
function displayCurrentData() {
  showDebugInfo(`显示当前数据，索引: ${currentIndex}`);
  
  try {
    // 检查模块是否已加载
    if (!DataVisualizer) {
      showDebugInfo('错误: DataVisualizer模块未加载');
      return;
    }
    
    if (!dataContainer) {
      const errorMsg = '数据容器不存在';
      showDebugInfo(errorMsg);
      console.error(errorMsg);
      return;
    }
    
    if (!currentData || currentData.length === 0) {
      dataContainer.innerHTML = '<p>没有数据可显示</p>';
      if (paginationControls) {
        paginationControls.innerHTML = '';
      }
      showDebugInfo('没有数据可显示');
      return;
    }
    
    // 确保索引在有效范围内
    if (currentIndex >= currentData.length) {
      currentIndex = currentData.length - 1;
    }
    if (currentIndex < 0) {
      currentIndex = 0;
    }
    
    // 获取当前数据
    const rowData = currentData[currentIndex];
    showDebugInfo(`当前行数据键数: ${rowData ? Object.keys(rowData).length : 0}`);
    
    // 根据当前视图显示数据
    showLoadingStatus('正在显示数据...');
    showDebugInfo(`准备显示视图: ${currentView}`);
    switch (currentView) {
      case 'single':
        showDebugInfo('调用createSingleDataView');
        DataVisualizer.createSingleDataView(rowData, dataContainer);
        break;
      case 'json':
        showDebugInfo('调用createJsonView');
        DataVisualizer.createJsonView(rowData, dataContainer);
        break;
      case 'table':
        showDebugInfo('调用createTable');
        // 仅显示当前行的表格
        DataVisualizer.createTable(currentData, dataContainer);
        break;
      default:
        showDebugInfo('调用默认createSingleDataView');
        DataVisualizer.createSingleDataView(rowData, dataContainer);
    }
    
    // 创建分页控件
    if (paginationControls) {
      showDebugInfo('创建分页控件');
      DataVisualizer.createPagination(
        currentIndex,
        currentData.length,
        handlePageChange,
        paginationControls
      );
    }
    
    showDebugInfo('数据展示完成');
  } catch (error) {
    const errorMsg = `显示数据时出错: ${error.message}\n堆栈信息: ${error.stack}`;
    showDebugInfo(errorMsg);
    console.error('显示数据时出错:', error);
    if (dataContainer) {
      dataContainer.innerHTML = `<p class="error">显示数据时出错: ${error.message}</p>
                                 <p>详细错误信息已在调试区域显示</p>`;
    }
  }
}

// 页面变化处理
function handlePageChange(newIndex) {
  showDebugInfo(`页面变化，新索引: ${newIndex}`);
  
  try {
    currentIndex = newIndex;
    displayCurrentData();
  } catch (error) {
    const errorMsg = `处理页面变化时出错: ${error.message}\n堆栈信息: ${error.stack}`;
    showDebugInfo(errorMsg);
    console.error('处理页面变化时出错:', error);
  }
}

// 显示数据信息
function displayDataInfo() {
  try {
    if (!dataInfo) {
      showDebugInfo('数据信息容器不存在');
      return;
    }
    
    const totalRows = originalData ? originalData.length : 0;
    const filteredRows = currentData ? currentData.length : 0;
    
    dataInfo.innerHTML = `
      <p>总共: ${totalRows} 行数据</p>
      ${filteredRows !== totalRows ? `<p>筛选后: ${filteredRows} 行数据</p>` : ''}
    `;
    showDebugInfo(`数据信息更新完成，总共: ${totalRows} 行，筛选后: ${filteredRows} 行`);
  } catch (error) {
    const errorMsg = `显示数据信息时出错: ${error.message}\n堆栈信息: ${error.stack}`;
    showDebugInfo(errorMsg);
    console.error('显示数据信息时出错:', error);
  }
}

// 页面加载完成后初始化应用
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', checkModulesLoaded);
} else {
  checkModulesLoaded();
}