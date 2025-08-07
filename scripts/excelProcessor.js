import * as XLSX from '../node_modules/xlsx/xlsx.mjs';

/**
 * 处理Excel数据，特别是处理换行符
 */
export class ExcelProcessor {
    /**
     * 读取Excel文件
     * @param {File} file - Excel文件
     * @returns {Promise} 解析后的数据
     */
    static readFile(file) {
        return new Promise((resolve, reject) => {
            try {
                if (!file) {
                    reject(new Error('文件对象为空'));
                    return;
                }
                
                const reader = new FileReader();
                
                reader.onload = (e) => {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array' });
                        
                        // 获取第一个工作表
                        const firstSheetName = workbook.SheetNames[0];
                        if (!firstSheetName) {
                            reject(new Error('Excel文件中没有工作表'));
                            return;
                        }
                        
                        const worksheet = workbook.Sheets[firstSheetName];
                        if (!worksheet) {
                            reject(new Error('无法读取工作表内容'));
                            return;
                        }
                        
                        // 将工作表转换为JSON
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        
                        resolve({ 
                            data: jsonData, 
                            workbook: workbook, 
                            worksheet: worksheet, 
                            sheetName: firstSheetName
                        });
                    } catch (error) {
                        reject(new Error('解析Excel文件时出错: ' + error.message));
                    }
                };
                
                reader.onerror = (error) => {
                    reject(new Error('读取文件时出错: ' + error.message));
                };
                
                reader.readAsArrayBuffer(file);
            } catch (error) {
                reject(new Error('处理文件时发生未知错误: ' + error.message));
            }
        });
    }
    
    /**
     * 处理数据，特别是处理第六列中的换行符
     * @param {Array} data - 原始数据
     * @returns {Array} 处理后的数据
     */
    static processData(data) {
        try {
            if (!data || data.length === 0) {
                return {
                    headers: [],
                    data: []
                };
            }
            
            // 处理表头
            const headers = data[0] || [];
            
            // 分批处理数据以提高性能
            const batchSize = 1000; // 每批处理1000行
            const processedData = [];
            
            // 分批处理数据行
            for (let i = 1; i < data.length; i += batchSize) {
                const batch = data.slice(i, Math.min(i + batchSize, data.length));
                const processedBatch = batch.map((row, index) => {
                    try {
                        // 处理第六列(索引为5)的换行符
                        if (row && row[5]) {
                            // 将换行符统一为\n
                            row[5] = row[5].toString().replace(/\r\n/g, '\n').replace(/\r/g, '\n');
                        }
                        
                        // 创建带有列名的对象
                        const processedRow = {};
                        headers.forEach((header, j) => {
                            processedRow[header] = row && row[j] !== undefined ? row[j] : '';
                        });
                        
                        // 添加行号（用于调试）
                        processedRow._rowIndex = i + index + 1; // +1 因为索引从0开始，且跳过了表头
                        
                        return processedRow;
                    } catch (error) {
                        console.warn('处理第' + (i + index + 1) + '行数据时出错:', error);
                        return {};
                    }
                }).filter(row => row && Object.keys(row).length > 0); // 过滤掉空行
                
                processedData.push(...processedBatch);
            }
            
            return {
                headers: headers,
                data: processedData
            };
        } catch (error) {
            console.error('处理数据时出错:', error);
            return {
                headers: [],
                data: []
            };
        }
    }
    
    /**
     * 根据多个维度筛选数据
     * @param {Array} data - 原始数据
     * @param {Object} filters - 筛选条件
     * @returns {Array} 筛选后的数据
     */
    static filterData(data, filters) {
        try {
            if (!data || data.length === 0) {
                return [];
            }
            
            if (!filters || Object.keys(filters).length === 0) {
                return data;
            }
            
            return data.filter(row => {
                try {
                    return Object.keys(filters).every(key => {
                        const filterValue = filters[key];
                        if (!filterValue || filterValue.length === 0) {
                            return true;
                        }
                        
                        const cellValue = row[key];
                        if (cellValue === undefined || cellValue === null) {
                            return false;
                        }
                        
                        // 支持字符串包含匹配
                        return cellValue.toString().toLowerCase().includes(filterValue.toLowerCase());
                    });
                } catch (error) {
                    console.warn('筛选数据时出错:', error);
                    return true; // 出错时不过滤该行
                }
            });
        } catch (error) {
            console.error('筛选数据时出错:', error);
            return data; // 出错时返回原数据
        }
    }
    
    /**
     * 获取某一列的所有唯一值
     * @param {Array} data - 数据
     * @param {string} columnName - 列名
     * @returns {Array} 唯一值列表
     */
    static getColumnUniqueValues(data, columnName) {
        try {
            if (!data || data.length === 0 || !columnName) {
                return [];
            }
            
            // 限制只取前1000个唯一值以提高性能
            const maxUniqueValues = 1000;
            const values = new Set();
            
            for (let i = 0; i < data.length && values.size < maxUniqueValues; i++) {
                const row = data[i];
                if (row && row[columnName] !== undefined && row[columnName] !== null) {
                    values.add(row[columnName]);
                }
            }
            
            // 转换为数组并排序
            return Array.from(values).sort();
        } catch (error) {
            console.error('获取列 ' + columnName + ' 的唯一值时出错:', error);
            return [];
        }
    }
    
    /**
     * 保存数据到Excel文件（直接修改原始文件）
     * @param {Object} workbookData - 包含工作簿信息的对象
     * @param {Array} headers - 表头
     * @param {Array} data - 数据
     * @param {string} filename - 文件名
     */
    static saveToFile(workbookData, headers, data, filename) {
        try {
            // 获取原始工作簿和工作表
            const workbook = workbookData.workbook;
            const sheetName = workbookData.sheetName;
            
            // 准备数据（将对象数组转换为二维数组）
            const worksheetData = [];
            
            // 添加表头
            worksheetData.push(headers);
            
            // 添加数据行
            data.forEach(row => {
                const rowData = [];
                headers.forEach(header => {
                    rowData.push(row[header] !== undefined ? row[header] : '');
                });
                worksheetData.push(rowData);
            });
            
            // 创建新的工作表
            const newWorksheet = XLSX.utils.aoa_to_sheet(worksheetData);
            
            // 替换原始工作表
            workbook.Sheets[sheetName] = newWorksheet;
            
            // 导出文件
            XLSX.writeFile(workbook, filename);
        } catch (error) {
            console.error('保存文件时出错:', error);
            throw new Error('保存文件时出错: ' + error.message);
        }
    }
    
    /**
     * 更新特定行的数据
     * @param {Object} workbookData - 包含工作簿信息的对象
     * @param {Array} headers - 表头
     * @param {Array} data - 所有数据
     * @param {Object} rowData - 要更新的行数据
     * @param {number} rowIndex - 行索引（从0开始，不包括表头）
     */
    static updateRowInPlace(workbookData, headers, data, rowData, rowIndex) {
        try {
            // 更新内存中的数据
            if (rowIndex >= 0 && rowIndex < data.length) {
                data[rowIndex] = { ...data[rowIndex], ...rowData };
                return true;
            }
            return false;
        } catch (error) {
            console.error('更新行数据时出错:', error);
            throw new Error('更新行数据时出错: ' + error.message);
        }
    }
}