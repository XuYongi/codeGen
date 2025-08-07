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
let saveButton, saveToOriginalButton; // 添加保存按钮引用
let filterToggle, toggleFilterButton; // 筛选项折叠相关元素
let floatingTagPanel, togglePanelButton, manualTagSelect, manualRemarkInput, saveTagButton; // 悬浮窗面板元素

// 全局变量
let originalData = [];
let currentData = [];
let headers = [];
let currentIndex = 0;
let currentView = 'single'; // 'single', 'json', 'table'
let originalFile = null; // 保存原始文件引用
let workbookData = null; // 保存工作簿数据
let isFilterExpanded = false; // 筛选项是否展开

// 显示调试信息的函数
function showDebugInfo(message) {
  if (debugInfo) {
    const time = new Date().toLocaleTimeString();
    debugInfo.innerHTML += '<p>[' + time + '] ' + message + '</p>';
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
    dataContainer.innerHTML = '<p>⏳ ' + message + '</p>';
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
  filterToggle = document.getElementById('filter-toggle');
  toggleFilterButton = document.getElementById('toggle-filter-button');
  
  // 获取悬浮窗面板元素
  floatingTagPanel = document.getElementById('floating-tag-panel');
  togglePanelButton = document.getElementById('toggle-panel');
  manualTagSelect = document.getElementById('manual-tag');
  manualRemarkInput = document.getElementById('manual-remark');
  saveTagButton = document.getElementById('save-tag');
  
  // 检查DOM元素是否存在
  if (!fileInput) {
    showDebugInfo('错误: 无法找到文件输入元素');
    return;
  }
  
  showDebugInfo('DOM元素获取完成');
  
  // 绑定事件监听器
  fileInput.addEventListener('change', handleFileSelect);
  
  // 绑定筛选项折叠按钮事件
  if (toggleFilterButton) {
    toggleFilterButton.addEventListener('click', toggleFilterDisplay);
  }
  
  // 绑定悬浮窗面板事件
  if (togglePanelButton) {
    togglePanelButton.addEventListener('click', toggleFloatingPanel);
  }
  
  if (saveTagButton) {
    saveTagButton.addEventListener('click', saveManualTagFromPanel);
  }
  
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
  
  // 保存原始文件引用
  originalFile = file;
  
  showDebugInfo('选择的文件: ' + file.name + ', 大小: ' + file.size + ' bytes');
  
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
    
    // 隐藏上传区域
    const uploadSection = document.getElementById('upload-section');
    if (uploadSection) {
      uploadSection.classList.add('hidden');
    }
    
    showDebugInfo('开始读取Excel文件');
    // 读取Excel文件
    const result = await ExcelProcessor.readFile(file);
    workbookData = result; // 保存工作簿数据
    showDebugInfo('原始数据读取完成，数据行数: ' + (result.data ? result.data.length : 0));
    
    // 处理数据
    showLoadingStatus('正在处理数据...');
    showDebugInfo('开始处理数据');
    const processedData = ExcelProcessor.processData(result.data);
    showDebugInfo('数据处理完成，处理后的数据条数: ' + (processedData.data ? processedData.data.length : 0));
    
    originalData = processedData.data || [];
    currentData = [...originalData];
    headers = processedData.headers || [];
    
    // 添加人工标签和人工备注列（如果不存在）
    if (!headers.includes('manualTag')) {
      headers.push('manualTag');
    }
    if (!headers.includes('manualRemark')) {
      headers.push('manualRemark');
    }
    
    showDebugInfo('数据赋值完成，原始数据条数: ' + originalData.length);
    
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
      
      // 初始化筛选控件折叠功能
      initializeFilterCollapse();
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
    
    // 添加保存按钮
    createSaveButtons();
    
    // 显示数据
    showLoadingStatus('正在显示数据...');
    showDebugInfo('显示数据信息');
    displayDataInfo();
    showDebugInfo('显示当前数据');
    displayCurrentData();
    
    showDebugInfo('文件处理完成');
  } catch (error) {
    const errorMsg = '处理文件时出错: ' + error.message + '\n堆栈信息: ' + error.stack;
    showDebugInfo(errorMsg);
    console.error('处理文件时出错:', error);
    if (dataContainer) {
      dataContainer.innerHTML = '<p class="error">处理文件时出错: ' + error.message + '</p>' +
                                 '<p>请检查文件格式是否正确（支持 .xlsx 和 .xls 格式）</p>' +
                                 '<p>详细错误信息已在调试区域显示</p>';
    }
  }
}

// 初始化筛选控件折叠功能
function initializeFilterCollapse() {
  if (!filterControls) return;
  
  const filterGroups = filterControls.querySelectorAll('.filter-group');
  if (filterGroups.length > 5) { // 如果筛选项超过5个，则启用折叠功能
    // 隐藏除前5个以外的筛选项
    for (let i = 5; i < filterGroups.length; i++) {
      filterGroups[i].classList.add('hidden-filter');
      filterGroups[i].style.display = 'none';
    }
    
    // 显示折叠按钮
    if (filterToggle) {
      filterToggle.classList.remove('hidden');
    }
    
    isFilterExpanded = false;
    updateFilterToggleButton();
  }
}

// 切换筛选项显示状态
function toggleFilterDisplay() {
  if (!filterControls) return;
  
  const filterGroups = filterControls.querySelectorAll('.filter-group');
  if (filterGroups.length <= 5) return; // 筛选项不超过5个，无需折叠
  
  isFilterExpanded = !isFilterExpanded;
  
  // 切换隐藏筛选项的显示状态
  for (let i = 5; i < filterGroups.length; i++) {
    if (isFilterExpanded) {
      filterGroups[i].style.display = 'flex';
    } else {
      filterGroups[i].style.display = 'none';
    }
  }
  
  updateFilterToggleButton();
}

// 更新筛选项折叠按钮文本
function updateFilterToggleButton() {
  if (!toggleFilterButton) return;
  
  toggleFilterButton.textContent = isFilterExpanded ? '收起筛选项' : '显示更多筛选项';
}

// 创建保存按钮
function createSaveButtons() {
  // 移除已存在的保存按钮
  if (saveButton) {
    saveButton.remove();
  }
  if (saveToOriginalButton) {
    saveToOriginalButton.remove();
  }
  
  // 创建保存按钮（保存为新文件）
  saveButton = document.createElement('button');
  saveButton.textContent = '保存为新文件';
  saveButton.style.marginLeft = '1rem';
  saveButton.style.padding = '0.5rem 1rem';
  saveButton.addEventListener('click', saveMarkedData);
  
  // 创建保存到原始文件按钮
  saveToOriginalButton = document.createElement('button');
  saveToOriginalButton.textContent = '保存到原始文件';
  saveToOriginalButton.style.marginLeft = '1rem';
  saveToOriginalButton.style.padding = '0.5rem 1rem';
  saveToOriginalButton.addEventListener('click', saveToOriginalFile);
  
  // 将保存按钮添加到视图控制区域
  if (viewControls) {
    viewControls.appendChild(saveButton);
    viewControls.appendChild(saveToOriginalButton);
  }
}

// 保存标记数据到新文件
async function saveMarkedData() {
  try {
    if (!ExcelProcessor) {
      alert('ExcelProcessor模块未加载');
      return;
    }
    
    if (!originalFile) {
      alert('未加载原始文件');
      return;
    }
    
    if (!workbookData || !headers || !currentData || currentData.length === 0) {
      alert('没有数据可保存');
      return;
    }
    
    showLoadingStatus('正在保存数据...');
    
    // 生成文件名
    const originalName = originalFile.name;
    const nameWithoutExt = originalName.substring(0, originalName.lastIndexOf('.'));
    const ext = originalName.substring(originalName.lastIndexOf('.'));
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const newFilename = nameWithoutExt + '_marked_' + timestamp + ext;
    
    // 保存文件
    ExcelProcessor.saveToFile(workbookData, headers, currentData, newFilename);
    
    showDebugInfo('文件保存成功: ' + newFilename);
    alert('文件保存成功: ' + newFilename);
    
    // 重新显示当前数据
    displayCurrentData();
  } catch (error) {
    const errorMsg = '保存文件时出错: ' + error.message + '\n堆栈信息: ' + error.stack;
    showDebugInfo(errorMsg);
    console.error('保存文件时出错:', error);
    alert('保存文件时出错: ' + error.message);
  }
}

// 保存标记数据到原始文件
async function saveToOriginalFile() {
  try {
    if (!ExcelProcessor) {
      alert('ExcelProcessor模块未加载');
      return;
    }
    
    if (!originalFile) {
      alert('未加载原始文件');
      return;
    }
    
    if (!workbookData || !headers || !currentData || currentData.length === 0) {
      alert('没有数据可保存');
      return;
    }
    
    // 确认操作
    if (!confirm('确定要直接修改原始文件吗？此操作不可撤销。')) {
      return;
    }
    
    showLoadingStatus('正在保存数据到原始文件...');
    
    // 保存文件（覆盖原始文件）
    ExcelProcessor.saveToFile(workbookData, headers, currentData, originalFile.name);
    
    showDebugInfo('原始文件已更新: ' + originalFile.name);
    alert('原始文件已更新: ' + originalFile.name);
    
    // 重新显示当前数据
    displayCurrentData();
  } catch (error) {
    const errorMsg = '保存文件时出错: ' + error.message + '\n堆栈信息: ' + error.stack;
    showDebugInfo(errorMsg);
    console.error('保存文件时出错:', error);
    alert('保存文件时出错: ' + error.message);
  }
}

// 筛选条件变化处理
function handleFilterChange(filters) {
  showDebugInfo('筛选条件变化: ' + JSON.stringify(filters));
  
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
      showDebugInfo('筛选完成，筛选前: ' + oldLength + ' 条，筛选后: ' + currentData.length + ' 条');
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
    const errorMsg = '处理筛选条件变化时出错: ' + error.message + '\n堆栈信息: ' + error.stack;
    showDebugInfo(errorMsg);
    console.error('处理筛选条件变化时出错:', error);
  }
}

// 视图变化处理
function handleViewChange(view) {
  showDebugInfo('视图变化: ' + view);
  
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
      
      // 重新添加保存按钮
      createSaveButtons();
    }
    
    // 重新显示当前数据
    displayCurrentData();
  } catch (error) {
    const errorMsg = '处理视图变化时出错: ' + error.message + '\n堆栈信息: ' + error.stack;
    showDebugInfo(errorMsg);
    console.error('处理视图变化时出错:', error);
  }
}

// 显示当前数据
function displayCurrentData() {
  showDebugInfo('显示当前数据，索引: ' + currentIndex);
  
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
    showDebugInfo('当前行数据键数: ' + (rowData ? Object.keys(rowData).length : 0));
    
    // 根据当前视图显示数据
    showLoadingStatus('正在显示数据...');
    showDebugInfo('准备显示视图: ' + currentView);
    switch (currentView) {
      case 'single':
        showDebugInfo('调用createSingleDataView');
        DataVisualizer.createSingleDataView(rowData, dataContainer, currentIndex);
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
        DataVisualizer.createSingleDataView(rowData, dataContainer, currentIndex);
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
    
    // 更新悬浮窗面板标记信息
    updateFloatingPanelTagInfo(rowData);
    
    // 显示悬浮窗面板
    showFloatingPanel();
    
    showDebugInfo('数据展示完成');
  } catch (error) {
    const errorMsg = '显示数据时出错: ' + error.message + '\n堆栈信息: ' + error.stack;
    showDebugInfo(errorMsg);
    console.error('显示数据时出错:', error);
    if (dataContainer) {
      dataContainer.innerHTML = '<p class="error">显示数据时出错: ' + error.message + '</p>' +
                                 '<p>详细错误信息已在调试区域显示</p>';
    }
  }
}

// 页面变化处理
function handlePageChange(newIndex) {
  showDebugInfo('页面变化，新索引: ' + newIndex);
  
  try {
    currentIndex = newIndex;
    displayCurrentData();
  } catch (error) {
    const errorMsg = '处理页面变化时出错: ' + error.message + '\n堆栈信息: ' + error.stack;
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
    
    dataInfo.innerHTML = '<p>总共: ' + totalRows + ' 行数据</p>' +
      (filteredRows !== totalRows ? '<p>筛选后: ' + filteredRows + ' 行数据</p>' : '');
    showDebugInfo('数据信息更新完成，总共: ' + totalRows + ' 行，筛选后: ' + filteredRows + ' 行');
  } catch (error) {
    const errorMsg = '显示数据信息时出错: ' + error.message + '\n堆栈信息: ' + error.stack;
    showDebugInfo(errorMsg);
    console.error('显示数据信息时出错:', error);
  }
}

// 更新悬浮窗面板标记信息
function updateFloatingPanelTagInfo(rowData) {
  if (!manualTagSelect || !manualRemarkInput) return;
  
  // 设置当前标签值
  if (rowData.manualTag) {
    manualTagSelect.value = rowData.manualTag;
  } else {
    manualTagSelect.value = '';
  }
  
  // 设置当前备注值
  if (rowData.manualRemark) {
    manualRemarkInput.value = rowData.manualRemark;
  } else {
    manualRemarkInput.value = '';
  }
}

// 显示悬浮窗面板
function showFloatingPanel() {
  if (floatingTagPanel) {
    floatingTagPanel.classList.remove('hidden');
  }
}

// 隐藏悬浮窗面板
function hideFloatingPanel() {
  if (floatingTagPanel) {
    floatingTagPanel.classList.add('hidden');
  }
}

// 切换悬浮窗面板显示状态（展开/收起）
function toggleFloatingPanel() {
  if (!floatingTagPanel || !togglePanelButton) return;
  
  const isCollapsed = floatingTagPanel.classList.contains('collapsed');
  
  if (isCollapsed) {
    // 展开面板
    floatingTagPanel.classList.remove('collapsed');
    togglePanelButton.textContent = '◀';
  } else {
    // 收起面板
    floatingTagPanel.classList.add('collapsed');
    togglePanelButton.textContent = '▶';
  }
}

// 从悬浮窗面板保存人工标记
function saveManualTagFromPanel() {
  if (!currentData || currentIndex >= currentData.length) {
    alert('没有选中的数据行');
    return;
  }
  
  const rowData = currentData[currentIndex];
  const tag = manualTagSelect ? manualTagSelect.value : '';
  const remark = manualRemarkInput ? manualRemarkInput.value : '';
  
  // 更新数据对象
  rowData.manualTag = tag;
  rowData.manualRemark = remark;
  
  // 显示保存成功提示
  alert('标记保存成功！');
  
  showDebugInfo('保存标记: ' + JSON.stringify({ rowIndex: currentIndex, tag, remark }));
}

// 页面加载完成后初始化应用
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', checkModulesLoaded);
} else {
  checkModulesLoaded();
}