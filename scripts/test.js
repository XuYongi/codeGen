// 简单测试脚本，用于验证Excel处理功能
console.log('测试脚本加载成功');

// 测试ExcelProcessor是否可以正常导入
import('./excelProcessor.js')
  .then(module => {
    console.log('ExcelProcessor模块导入成功');
    console.log('ExcelProcessor:', module.ExcelProcessor);
  })
  .catch(error => {
    console.error('ExcelProcessor模块导入失败:', error);
  });

// 测试DataVisualizer是否可以正常导入
import('./dataVisualizer.js')
  .then(module => {
    console.log('DataVisualizer模块导入成功');
    console.log('DataVisualizer:', module.DataVisualizer);
  })
  .catch(error => {
    console.error('DataVisualizer模块导入失败:', error);
  });