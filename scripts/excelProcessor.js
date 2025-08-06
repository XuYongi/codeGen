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
                        
                        resolve(jsonData);
                    } catch (error) {
                        reject(new Error(`解析Excel文件时出错: ${error.message}`));
                    }
                };
                
                reader.onerror = (error) => {
                    reject(new Error(`读取文件时出错: ${error.message}`));
                };
                
                reader.readAsArrayBuffer(file);
            } catch (error) {
                reject(new Error(`处理文件时发生未知错误: ${error.message}`));
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
            
            // 处理数据行
            const processedData = data.slice(1).map((row, index) => {
                try {
                    // 处理第六列(索引为5)的换行符
                    if (row && row[5]) {
                        // 将换行符统一为\n
                        row[5] = row[5].toString().replace(/\r\n/g, '\n').replace(/\r/g, '\n');
                    }
                    
                    // 创建带有列名的对象
                    const processedRow = {};
                    headers.forEach((header, i) => {
                        processedRow[header] = row && row[i] !== undefined ? row[i] : '';
                    });
                    
                    // 添加行号（用于调试）
                    processedRow._rowIndex = index + 2; // +2 因为索引从0开始，且跳过了表头
                    
                    return processedRow;
                } catch (error) {
                    console.warn(`处理第${index + 2}行数据时出错:`, error);
                    return {};
                }
            }).filter(row => row && Object.keys(row).length > 0); // 过滤掉空行
            
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
            
            const values = data.map(row => {
                if (row && row[columnName] !== undefined && row[columnName] !== null) {
                    return row[columnName];
                }
                return '';
            }).filter(value => value !== '');
            
            // 去重并排序
            return [...new Set(values)].sort();
        } catch (error) {
            console.error(`获取列 ${columnName} 的唯一值时出错:`, error);
            return [];
        }
    }
}