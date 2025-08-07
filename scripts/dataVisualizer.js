import { ExcelProcessor } from './excelProcessor.js';

/**
 * 数据可视化组件
 */
export class DataVisualizer {
    /**
     * 创建数据表格（只显示指定列）
     * @param {Array} data - 要显示的数据
     * @param {HTMLElement} container - 容器元素
     */
    static createTable(data, container) {
        if (!container) {
            console.error('容器元素未指定');
            return;
        }
        
        try {
            // 清空容器
            container.innerHTML = '';
            
            if (!data || data.length === 0) {
                container.innerHTML = '<p>没有数据可显示</p>';
                return;
            }
            
            // 限制显示的数据量以提高性能，只显示前100条
            const displayData = data.slice(0, 100);
            if (data.length > 100) {
                const warning = document.createElement('div');
                warning.className = 'warning';
                warning.innerHTML = '<p>注意：数据量较大，仅显示前100条记录。总共 ' + data.length + ' 条记录。</p>';
                container.appendChild(warning);
            }
            
            // 创建表格元素
            const table = document.createElement('table');
            table.className = 'data-table';
            
            // 创建表头（只显示指定列）
            const thead = document.createElement('thead');
            const headerRow = document.createElement('tr');
            
            // 指定要显示的列
            const displayColumns = [
                'model_output',
                'ide_codemate_model_feedback.accept_content',
                'accept_type',
                'prompt_sections'  // 拆分后的prompt字段
            ];
            
            const displayColumnHeaders = [
                'model_output',
                'accept_content',
                'accept_type',
                'prompt 拆分内容'
            ];
            
            displayColumnHeaders.forEach(headerText => {
                const th = document.createElement('th');
                th.textContent = headerText;
                headerRow.appendChild(th);
            });
            
            thead.appendChild(headerRow);
            table.appendChild(thead);
            
            // 创建表体
            const tbody = document.createElement('tbody');
            
            displayData.forEach((row, rowIndex) => {
                const tr = document.createElement('tr');
                
                // 显示指定的列
                displayColumns.forEach((column, index) => {
                    const td = document.createElement('td');
                    
                    if (column === 'prompt_sections') {
                        // 从extra_params中提取并拆分prompt字段
                        const extraParamsStr = row['ide_codemate_model_request.extra_params'];
                        console.log('表格模式 - extra_params内容:', extraParamsStr);
                        
                        if (extraParamsStr) {
                            try {
                                // 如果extra_params是JSON字符串，则解析并提取prompt
                                if (typeof extraParamsStr === 'string' && 
                                    (extraParamsStr.trim().startsWith('{') || extraParamsStr.trim().startsWith('['))) {
                                    const extraParams = JSON.parse(extraParamsStr);
                                    const promptValue = extraParams.prompt || '';
                                    console.log('表格模式 - 解析出的prompt:', promptValue);
                                    
                                    // 拆分prompt内容
                                    const sections = this.parsePromptSections(promptValue);
                                    console.log('表格模式 - 拆分后的sections:', sections);
                                    
                                    // 创建展示容器
                                    const sectionsContainer = document.createElement('div');
                                    sectionsContainer.className = 'prompt-sections';
                                    
                                    // 检查是否有拆分内容
                                    if (Object.keys(sections).length === 0) {
                                        sectionsContainer.textContent = '无拆分内容';
                                    } else {
                                        // 为每个部分创建展示元素，但限制显示的部分数量
                                        const sectionKeys = Object.keys(sections);
                                        const maxSections = 5; // 限制最多显示5个部分
                                        
                                        sectionKeys.slice(0, maxSections).forEach(sectionKey => {
                                            const sectionDiv = document.createElement('div');
                                            sectionDiv.className = 'prompt-section';
                                            
                                            const sectionTitle = document.createElement('div');
                                            sectionTitle.className = 'prompt-section-title';
                                            sectionTitle.textContent = sectionKey;
                                            
                                            const sectionContent = document.createElement('pre');
                                            sectionContent.className = 'code-block language-java';
                                            // 设置默认显示15行高度，并添加展开/折叠功能
                                            sectionContent.style.height = '150px'; // 表格中默认显示约15行
                                            sectionContent.style.overflow = 'auto'; // 添加滚动条
                                            sectionContent.style.position = 'relative';
                                            sectionContent.style.resize = 'vertical';
                                            // 保持原始格式，包括缩进
                                            sectionContent.textContent = sections[sectionKey];
                                            
                                            // 添加展开/折叠按钮
                                            const toggleButton = document.createElement('button');
                                            toggleButton.className = 'code-toggle-button';
                                            toggleButton.textContent = '展开';
                                            toggleButton.style.position = 'absolute';
                                            toggleButton.style.bottom = '10px';
                                            toggleButton.style.right = '10px';
                                            toggleButton.style.padding = '5px 10px';
                                            toggleButton.style.background = '#3498db';
                                            toggleButton.style.color = 'white';
                                            toggleButton.style.border = 'none';
                                            toggleButton.style.borderRadius = '4px';
                                            toggleButton.style.cursor = 'pointer';
                                            
                                            toggleButton.addEventListener('click', function() {
                                                if (sectionContent.style.height === '150px' || sectionContent.style.height === '') {
                                                    sectionContent.style.height = 'auto'; // 展开到完整内容
                                                    toggleButton.textContent = '折叠';
                                                    
                                                    // 重新应用语法高亮
                                                    if (typeof Prism !== 'undefined') {
                                                        Prism.highlightElement(sectionContent);
                                                    }
                                                } else {
                                                    sectionContent.style.height = '150px';
                                                    toggleButton.textContent = '展开';
                                                }
                                            });
                                            
                                            sectionContent.appendChild(toggleButton);
                                            sectionDiv.appendChild(sectionTitle);
                                            sectionDiv.appendChild(sectionContent);
                                            sectionsContainer.appendChild(sectionDiv);
                                            
                                            // 应用语法高亮
                                            if (typeof Prism !== 'undefined') {
                                                Prism.highlightElement(sectionContent);
                                            }
                                        });
                                        
                                        // 如果有更多部分，显示提示
                                        if (sectionKeys.length > maxSections) {
                                            const moreDiv = document.createElement('div');
                                            moreDiv.className = 'more-sections';
                                            moreDiv.textContent = '... 还有 ' + (sectionKeys.length - maxSections) + ' 个部分';
                                            sectionsContainer.appendChild(moreDiv);
                                        }
                                    }
                                    
                                    td.appendChild(sectionsContainer);
                                } else {
                                    // 限制显示的字符数以提高性能
                                    let displayText = extraParamsStr;
                                    if (displayText && displayText.length > 70) {
                                        displayText = displayText.substring(0, 70) + '...';
                                    }
                                    td.textContent = displayText;
                                }
                            } catch (e) {
                                // 解析失败则显示原始值
                                console.warn('解析extra_params失败:', e);
                                td.textContent = '解析失败: ' + extraParamsStr;
                            }
                        } else {
                            td.textContent = '无数据';
                        }
                    } else {
                        // 其他列直接显示值，但限制显示的字符数
                        const cellValue = row[column];
                        if (cellValue !== undefined && cellValue !== null) {
                            let displayText = cellValue.toString();
                            if (displayText.length > 100) {
                                displayText = displayText.substring(0, 100) + '...';
                            }
                            td.textContent = displayText;
                        } else {
                            td.textContent = '无数据';
                        }
                    }
                    
                    tr.appendChild(td);
                });
                
                tbody.appendChild(tr);
            });
            
            table.appendChild(tbody);
            container.appendChild(table);
            
            // 应用语法高亮
            this.applySyntaxHighlighting(container);
        } catch (error) {
            console.error('创建表格时出错:', error);
            container.innerHTML = '<p class="error">创建表格时出错: ' + error.message + '</p>';
        }
    }
    
    /**
     * 解析并拆分prompt内容
     * @param {string} promptText - 完整的prompt文本
     * @returns {Object} 拆分后的内容对象
     */
    static parsePromptSections(promptText) {
        const sections = {};
        
        if (!promptText) {
            return sections;
        }
        
        try {
            console.log('开始解析prompt内容:', promptText.substring(0, 100) + '...');
            
            // 特殊处理，尝试识别常见的结构
            if (promptText.includes('Below is the code context:')) {
                console.log('检测到特定格式的prompt');
                const contextStart = promptText.indexOf('Below is the code context:');
                const snippetStart = promptText.indexOf('And here is the code snippet you are asked to complete:');
                
                if (contextStart !== -1) {
                    sections['任务描述'] = promptText.substring(0, contextStart).trim();
                    console.log('任务描述:', sections['任务描述'].substring(0, 100) + '...');
                    
                    if (snippetStart !== -1) {
                        sections['Below is the code context:'] = promptText.substring(contextStart + 'Below is the code context:'.length, snippetStart).trim();
                        sections['And here is the code snippet you are asked to complete:'] = promptText.substring(snippetStart + 'And here is the code snippet you are asked to complete:'.length).trim();
                        console.log('上下文内容:', sections['Below is the code context:'].substring(0, 100) + '...');
                        console.log('待完成代码:', sections['And here is the code snippet you are asked to complete:'].substring(0, 100) + '...');
                    } else {
                        sections['Below is the code context:'] = promptText.substring(contextStart + 'Below is the code context:'.length).trim();
                        console.log('上下文内容:', sections['Below is the code context:'].substring(0, 100) + '...');
                    }
                } else {
                    sections['默认内容'] = promptText;
                    console.log('默认内容:', sections['默认内容'].substring(0, 100) + '...');
                }
                console.log('特定格式解析完成，sections:', sections);
                return sections;
            }
            
            // 首先尝试按照标题模式拆分
            // 匹配以 # 开头的标题或者以 : 结尾的行作为标题
            const lines = promptText.split('\n');
            let currentSection = '';
            let currentContent = [];
            
            lines.forEach((line, index) => {
                // 检查是否为标题行
                if (line.match(/^#{1,6}\s+.+/) || line.match(/.+:\s*$/)) {
                    console.log('找到标题行:', line);
                    // 如果已有内容，则保存前一个部分
                    if (currentSection && currentContent.length > 0) {
                        sections[currentSection] = currentContent.join('\n').trim();
                        console.log('保存部分:', currentSection, '内容长度:', sections[currentSection].length);
                    }
                    
                    // 开始新的部分
                    currentSection = line.trim();
                    currentContent = [];
                    console.log('开始新部分:', currentSection);
                } else if (currentSection) {
                    // 添加内容到当前部分，但限制内容长度以提高性能
                    if (currentContent.join('\n').length < 5000) {
                        currentContent.push(line);
                    }
                }
            });
            
            // 保存最后一个部分
            if (currentSection && currentContent.length > 0) {
                sections[currentSection] = currentContent.join('\n').trim();
                console.log('保存最后部分:', currentSection, '内容长度:', sections[currentSection].length);
            }
            
            // 如果没有找到任何标题，则尝试其他方式拆分
            if (Object.keys(sections).length === 0) {
                console.log('未找到标题，尝试按段落拆分');
                // 尝试按段落拆分（以两个换行符为分隔）
                const paragraphs = promptText.split('\n\n');
                if (paragraphs.length > 1) {
                    // 限制段落数量以提高性能
                    const maxParagraphs = 10;
                    for (let i = 0; i < Math.min(paragraphs.length, maxParagraphs); i++) {
                        const paragraph = paragraphs[i];
                        if (paragraph.trim()) {
                            sections['段落 ' + (i + 1)] = paragraph.trim();
                            console.log('段落 ' + (i + 1) + ':', sections['段落 ' + (i + 1)].substring(0, 100) + '...');
                        }
                    }
                } else {
                    // 如果还是无法拆分，则将整个内容作为一项
                    sections['完整内容'] = promptText;
                    console.log('完整内容:', sections['完整内容'].substring(0, 100) + '...');
                }
            }
            
            console.log('解析完成，sections:', sections);
        } catch (error) {
            console.error('解析Prompt内容时出错:', error);
            sections['解析错误'] = promptText;
        }
        
        return sections;
    }
    
    /**
     * 创建单条数据展示视图
     * @param {Object} rowData - 单行数据
     * @param {HTMLElement} container - 容器元素
     * @param {number} rowIndex - 行索引
     */
    static createSingleDataView(rowData, container, rowIndex) {
        if (!container) {
            console.error('容器元素未指定');
            return;
        }
        
        try {
            // 清空容器
            container.innerHTML = '';
            
            if (!rowData) {
                container.innerHTML = '<p>没有数据可显示</p>';
                return;
            }
            
            // 创建数据展示容器
            const dataView = document.createElement('div');
            dataView.className = 'single-data-view';
            
            // 创建标题
            const title = document.createElement('h3');
            title.textContent = '数据详情';
            dataView.appendChild(title);
            
            // 创建数据展示区域
            const dataList = document.createElement('div');
            dataList.className = 'data-list';
            
            // 指定要显示的字段（按要求的顺序）
            const displayFields = [
                { key: 'model_output', label: 'model_output' },
                { key: 'ide_codemate_model_feedback.accept_content', label: 'accept_content' },
                { key: 'accept_type', label: 'accept_type' }
                // prompt部分单独处理，不放在这里
            ];
            
            // 为每个字段创建展示项（除了prompt）
            displayFields.forEach(field => {
                const item = document.createElement('div');
                item.className = 'data-item';
                
                const label = document.createElement('div');
                label.className = 'data-label';
                label.textContent = field.label;
                
                const value = document.createElement('div');
                value.className = 'data-value';
                
                const cellValue = rowData[field.key];
                
                // 处理包含换行符的字段，但限制显示的字符数
                if (typeof cellValue === 'string' && cellValue.includes('\n')) {
                    let displayText = cellValue;
                    if (displayText.length > 1000) {
                        displayText = displayText.substring(0, 1000) + '... (内容已截断)';
                    }
                    value.innerHTML = displayText.replace(/\n/g, '<br>');
                } else if (cellValue !== undefined && cellValue !== null) {
                    let displayText = cellValue.toString();
                    if (displayText.length > 1000) {
                        displayText = displayText.substring(0, 1000) + '... (内容已截断)';
                    }
                    value.textContent = displayText;
                } else {
                    value.textContent = '无数据';
                }
                
                item.appendChild(label);
                item.appendChild(value);
                dataList.appendChild(item);
            });
            
            // 单独处理prompt部分，让它独占一列
            const promptItem = document.createElement('div');
            promptItem.className = 'data-item';
            
            const promptLabel = document.createElement('div');
            promptLabel.className = 'data-label';
            promptLabel.textContent = 'prompt 拆分内容';
            
            const promptValue = document.createElement('div');
            promptValue.className = 'data-value';
            
            // 处理拆分后的prompt字段
            const extraParamsStr = rowData['ide_codemate_model_request.extra_params'];
            console.log('单条数据模式 - extra_params内容:', extraParamsStr);
            
            if (extraParamsStr) {
                try {
                    // 如果extra_params是JSON字符串，则解析并提取prompt
                    if (typeof extraParamsStr === 'string' && 
                        (extraParamsStr.trim().startsWith('{') || extraParamsStr.trim().startsWith('['))) {
                        const extraParams = JSON.parse(extraParamsStr);
                        const promptValueStr = extraParams.prompt || '';
                        console.log('单条数据模式 - 解析出的prompt:', promptValueStr);
                        
                        // 拆分prompt内容
                        const sections = this.parsePromptSections(promptValueStr);
                        console.log('单条数据模式 - 拆分后的sections:', sections);
                        
                        // 创建展示容器
                        const sectionsContainer = document.createElement('div');
                        sectionsContainer.className = 'prompt-sections';
                        
                        // 检查是否有拆分内容
                        if (Object.keys(sections).length === 0) {
                            const noContent = document.createElement('div');
                            noContent.textContent = '无拆分内容';
                            sectionsContainer.appendChild(noContent);
                        } else {
                            // 为每个部分创建展示元素
                            Object.keys(sections).forEach(sectionKey => {
                                const sectionDiv = document.createElement('div');
                                sectionDiv.className = 'prompt-section';
                                
                                const sectionTitle = document.createElement('div');
                                sectionTitle.className = 'prompt-section-title';
                                sectionTitle.textContent = sectionKey;
                                
                                const sectionContent = document.createElement('pre');
                                sectionContent.className = 'code-block language-java';
                                // 设置默认显示15行高度，并添加展开/折叠功能
                                sectionContent.style.height = '250px'; // 约15行的高度
                                sectionContent.style.overflow = 'auto'; // 添加滚动条
                                sectionContent.style.position = 'relative';
                                sectionContent.style.resize = 'vertical';
                                // 保持原始格式，包括缩进
                                sectionContent.textContent = sections[sectionKey];
                                
                                // 添加展开/折叠按钮
                                const toggleButton = document.createElement('button');
                                toggleButton.className = 'code-toggle-button';
                                toggleButton.textContent = '展开';
                                toggleButton.style.position = 'absolute';
                                toggleButton.style.bottom = '10px';
                                toggleButton.style.right = '10px';
                                toggleButton.style.padding = '5px 10px';
                                toggleButton.style.background = '#3498db';
                                toggleButton.style.color = 'white';
                                toggleButton.style.border = 'none';
                                toggleButton.style.borderRadius = '4px';
                                toggleButton.style.cursor = 'pointer';
                                
                                toggleButton.addEventListener('click', function() {
                                    if (sectionContent.style.height === '250px' || sectionContent.style.height === '') {
                                        sectionContent.style.height = 'auto'; // 展开到完整内容
                                        toggleButton.textContent = '折叠';
                                        
                                        // 重新应用语法高亮
                                        if (typeof Prism !== 'undefined') {
                                            Prism.highlightElement(sectionContent);
                                        }
                                    } else {
                                        sectionContent.style.height = '250px';
                                        toggleButton.textContent = '展开';
                                    }
                                });
                                
                                sectionContent.appendChild(toggleButton);
                                sectionDiv.appendChild(sectionTitle);
                                sectionDiv.appendChild(sectionContent);
                                sectionsContainer.appendChild(sectionDiv);
                                
                                // 应用语法高亮
                                if (typeof Prism !== 'undefined') {
                                    Prism.highlightElement(sectionContent);
                                }
                            });
                        }
                        
                        promptValue.appendChild(sectionsContainer);
                    } else {
                        const pre = document.createElement('pre');
                        pre.className = 'code-block language-java';
                        pre.textContent = '非JSON格式: ' + extraParamsStr || '';
                        promptValue.appendChild(pre);
                    }
                } catch (e) {
                    // 解析失败则显示原始值
                    console.warn('解析extra_params失败:', e);
                    const pre = document.createElement('pre');
                    pre.className = 'code-block language-java';
                    pre.textContent = '解析失败: ' + extraParamsStr || '';
                    promptValue.appendChild(pre);
                }
            } else {
                const noData = document.createElement('div');
                noData.textContent = '无数据';
                promptValue.appendChild(noData);
            }
            
            promptItem.appendChild(promptLabel);
            promptItem.appendChild(promptValue);
            dataList.appendChild(promptItem);
            
            // 显示其他字段（包括原来的extra_params）
            const otherFieldsToggle = document.createElement('div');
            otherFieldsToggle.className = 'other-fields-toggle';
            
            const toggleButton = document.createElement('button');
            toggleButton.textContent = '显示其他字段';
            toggleButton.onclick = () => this.toggleOtherFields(rowData, [...displayFields, { key: 'prompt_sections', label: 'prompt 拆分内容' }], dataList);
            
            otherFieldsToggle.appendChild(toggleButton);
            dataView.appendChild(dataList);
            dataView.appendChild(otherFieldsToggle);
            container.appendChild(dataView);
            
            // 应用语法高亮
            this.applySyntaxHighlighting(container);
        } catch (error) {
            console.error('创建单条数据视图时出错:', error);
            container.innerHTML = '<p class="error">创建单条数据视图时出错: ' + error.message + '</p>';
        }
    }
    
    /**
     * 切换显示其他字段
     * @param {Object} rowData - 单行数据
     * @param {Array} displayedFields - 已显示的字段
     * @param {HTMLElement} container - 容器元素
     */
    static toggleOtherFields(rowData, displayedFields, container) {
        try {
            // 获取已显示字段的键名
            const displayedKeys = displayedFields.map(f => f.key);
            
            // 查找未显示的字段（排除我们已经单独提取的字段）
            const excludeKeys = ['prompt_sections']; // 不在"其他字段"中重复显示的字段
            const otherFields = Object.keys(rowData)
                .filter(key => !displayedKeys.includes(key) && !excludeKeys.includes(key) && key !== '_rowIndex')
                .map(key => ({ key, label: key }));
            
            if (otherFields.length === 0) {
                return;
            }
            
            // 创建其他字段展示区域
            const otherFieldsContainer = document.createElement('div');
            otherFieldsContainer.className = 'other-fields-container';
            
            const title = document.createElement('h4');
            title.textContent = '其他字段';
            otherFieldsContainer.appendChild(title);
            
            const otherDataList = document.createElement('div');
            otherDataList.className = 'data-list';
            
            otherFields.forEach(field => {
                const item = document.createElement('div');
                item.className = 'data-item';
                
                const label = document.createElement('div');
                label.className = 'data-label';
                label.textContent = field.label;
                
                const value = document.createElement('div');
                value.className = 'data-value';
                
                const cellValue = rowData[field.key];
                
                // 特殊处理extra_params，以JSON格式展示
                if (field.key === 'ide_codemate_model_request.extra_params' && cellValue) {
                    try {
                        // 如果是字符串且看起来像JSON，则格式化显示
                        if (typeof cellValue === 'string' && 
                            (cellValue.trim().startsWith('{') || cellValue.trim().startsWith('['))) {
                            const parsed = JSON.parse(cellValue);
                            value.innerHTML = '<pre>' + JSON.stringify(parsed, null, 2) + '</pre>';
                        } else {
                            // 否则直接显示，但限制显示的字符数
                            let displayText = cellValue.toString();
                            if (displayText.length > 1000) {
                                displayText = displayText.substring(0, 1000) + '... (内容已截断)';
                            }
                            value.textContent = displayText;
                        }
                    } catch (e) {
                        // 解析失败则直接显示原始值，但限制显示的字符数
                        console.warn('解析extra_params失败:', e);
                        let displayText = cellValue.toString();
                        if (displayText.length > 1000) {
                            displayText = displayText.substring(0, 1000) + '... (内容已截断)';
                        }
                        value.textContent = displayText;
                    }
                } else if (typeof cellValue === 'string' && cellValue.includes('\n')) {
                    // 处理包含换行符的字段，但限制显示的字符数
                    let displayText = cellValue;
                    if (displayText.length > 1000) {
                        displayText = displayText.substring(0, 1000) + '... (内容已截断)';
                    }
                    value.innerHTML = displayText.replace(/\n/g, '<br>');
                } else if (cellValue !== undefined && cellValue !== null) {
                    // 限制显示的字符数
                    let displayText = cellValue.toString();
                    if (displayText.length > 1000) {
                        displayText = displayText.substring(0, 1000) + '... (内容已截断)';
                    }
                    value.textContent = displayText;
                } else {
                    value.textContent = '无数据';
                }
                
                item.appendChild(label);
                item.appendChild(value);
                otherDataList.appendChild(item);
            });
            
            otherFieldsContainer.appendChild(otherDataList);
            container.appendChild(otherFieldsContainer);
            
            // 移除切换按钮
            const toggleButton = container.parentElement.querySelector('.other-fields-toggle button');
            if (toggleButton) {
                toggleButton.parentElement.remove();
            }
            
            // 应用语法高亮
            this.applySyntaxHighlighting(container);
        } catch (error) {
            console.error('切换其他字段显示时出错:', error);
        }
    }
    
    /**
     * 创建JSON模式展示视图
     * @param {Object} rowData - 单行数据
     * @param {HTMLElement} container - 容器元素
     */
    static createJsonView(rowData, container) {
        if (!container) {
            console.error('容器元素未指定');
            return;
        }
        
        try {
            // 清空容器
            container.innerHTML = '';
            
            if (!rowData) {
                container.innerHTML = '<p>没有数据可显示</p>';
                return;
            }
            
            // 创建JSON展示容器
            const jsonView = document.createElement('div');
            jsonView.className = 'json-view';
            
            // 创建标题和控制按钮
            const header = document.createElement('div');
            header.className = 'json-header';
            
            const title = document.createElement('h3');
            title.textContent = 'JSON 数据';
            
            const toggleButton = document.createElement('button');
            toggleButton.className = 'toggle-button';
            toggleButton.textContent = '折叠';
            toggleButton.onclick = () => this.toggleJsonView(jsonContent);
            
            header.appendChild(title);
            header.appendChild(toggleButton);
            jsonView.appendChild(header);
            
            // 创建JSON内容区域
            const jsonContent = document.createElement('div');
            jsonContent.className = 'json-content expanded';
            
            // 格式化JSON数据，但限制长度以提高性能
            let formattedJson = JSON.stringify(rowData, null, 2);
            if (formattedJson.length > 50000) {
                formattedJson = formattedJson.substring(0, 50000) + '\n\n... (JSON内容已截断)';
            }
            const pre = document.createElement('pre');
            pre.textContent = formattedJson;
            jsonContent.appendChild(pre);
            
            jsonView.appendChild(jsonContent);
            container.appendChild(jsonView);
        } catch (error) {
            console.error('创建JSON视图时出错:', error);
            container.innerHTML = '<p class="error">创建JSON视图时出错: ' + error.message + '</p>';
        }
    }
    
    /**
     * 切换JSON视图的展开/折叠状态
     * @param {HTMLElement} contentElement - JSON内容元素
     */
    static toggleJsonView(contentElement) {
        try {
            if (!contentElement) return;
            
            const isExpanded = contentElement.classList.contains('expanded');
            const toggleButton = contentElement.parentElement.querySelector('.toggle-button');
            
            if (isExpanded) {
                contentElement.classList.remove('expanded');
                contentElement.classList.add('collapsed');
                if (toggleButton) {
                    toggleButton.textContent = '展开';
                }
            } else {
                contentElement.classList.remove('collapsed');
                contentElement.classList.add('expanded');
                if (toggleButton) {
                    toggleButton.textContent = '折叠';
                }
            }
        } catch (error) {
            console.error('切换JSON视图时出错:', error);
        }
    }
    
    /**
     * 创建分页控件
     * @param {number} currentIndex - 当前索引
     * @param {number} total - 总数
     * @param {Function} onPageChange - 页面变化回调
     * @param {HTMLElement} container - 容器元素
     */
    static createPagination(currentIndex, total, onPageChange, container) {
        if (!container) {
            console.error('容器元素未指定');
            return;
        }
        
        try {
            // 清空容器
            container.innerHTML = '';
            
            if (total <= 0) {
                return;
            }
            
            // 创建分页控件容器
            const pagination = document.createElement('div');
            pagination.className = 'pagination';
            
            // 创建上一页按钮
            const prevButton = document.createElement('button');
            prevButton.textContent = '上一条';
            prevButton.disabled = currentIndex <= 0;
            prevButton.onclick = () => {
                if (onPageChange) onPageChange(currentIndex - 1);
            };
            
            // 创建页码信息
            const pageInfo = document.createElement('span');
            pageInfo.className = 'page-info';
            pageInfo.textContent = (currentIndex + 1) + ' / ' + total;
            
            // 创建下一页按钮
            const nextButton = document.createElement('button');
            nextButton.textContent = '下一条';
            nextButton.disabled = currentIndex >= total - 1;
            nextButton.onclick = () => {
                if (onPageChange) onPageChange(currentIndex + 1);
            };
            
            // 添加页码输入和跳转功能
            const pageInput = document.createElement('input');
            pageInput.type = 'number';
            pageInput.min = 1;
            pageInput.max = total;
            pageInput.value = currentIndex + 1;
            pageInput.style.width = '60px';
            pageInput.style.margin = '0 5px';
            
            const goButton = document.createElement('button');
            goButton.textContent = '跳转';
            goButton.onclick = () => {
                const pageNumber = parseInt(pageInput.value);
                if (!isNaN(pageNumber) && pageNumber >= 1 && pageNumber <= total) {
                    if (onPageChange) onPageChange(pageNumber - 1);
                } else {
                    alert('请输入有效的页码 (1-' + total + ')');
                }
            };
            
            // 添加回车键支持
            pageInput.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') {
                    goButton.click();
                }
            });
            
            pagination.appendChild(prevButton);
            pagination.appendChild(pageInfo);
            pagination.appendChild(nextButton);
            pagination.appendChild(document.createTextNode(' 跳转到:'));
            pagination.appendChild(pageInput);
            pagination.appendChild(goButton);
            container.appendChild(pagination);
        } catch (error) {
            console.error('创建分页控件时出错:', error);
        }
    }
    
    /**
     * 创建视图切换控件
     * @param {string} currentView - 当前视图类型 ('single' | 'json' | 'table')
     * @param {Function} onViewChange - 视图变化回调
     * @param {HTMLElement} container - 容器元素
     */
    static createViewControls(currentView, onViewChange, container) {
        if (!container) {
            console.error('容器元素未指定');
            return;
        }
        
        try {
            // 清空容器
            container.innerHTML = '';
            
            // 创建视图控件容器
            const viewControls = document.createElement('div');
            viewControls.className = 'view-controls';
            
            // 创建标题
            const title = document.createElement('h3');
            title.textContent = '视图模式';
            viewControls.appendChild(title);
            
            // 创建视图选项按钮
            const views = [
                { id: 'single', label: '单条数据' },
                { id: 'json', label: 'JSON模式' },
                { id: 'table', label: '表格模式' }
            ];
            
            views.forEach(view => {
                const button = document.createElement('button');
                button.textContent = view.label;
                button.className = currentView === view.id ? 'active' : '';
                button.onclick = () => {
                    if (onViewChange) onViewChange(view.id);
                };
                viewControls.appendChild(button);
            });
            
            container.appendChild(viewControls);
        } catch (error) {
            console.error('创建视图控制时出错:', error);
        }
    }
    
    /**
     * 创建筛选控件（下拉框形式）
     * @param {Array} headers - 表头
     * @param {Array} data - 数据
     * @param {HTMLElement} container - 容器元素
     * @param {Function} onFilterChange - 筛选条件变化回调
     */
    static createFilterControls(headers, data, container, onFilterChange) {
        if (!container) {
            console.error('容器元素未指定');
            return;
        }
        
        try {
            // 清空容器
            container.innerHTML = '';
            
            if (!headers || headers.length === 0) {
                return;
            }
            
            // 创建筛选控件容器
            const filterContainer = document.createElement('div');
            filterContainer.className = 'filter-container';
            
            // 创建标题
            const title = document.createElement('h3');
            title.textContent = '筛选条件';
            filterContainer.appendChild(title);
            
            // 为每列创建筛选下拉框
            const filters = {};
            
            headers.filter(header => header !== '_rowIndex').forEach(header => {
                const filterGroup = document.createElement('div');
                filterGroup.className = 'filter-group';
                
                const label = document.createElement('label');
                label.textContent = header;
                label.setAttribute('for', 'filter-' + header);
                
                const select = document.createElement('select');
                select.id = 'filter-' + header;
                
                // 添加默认选项
                const defaultOption = document.createElement('option');
                defaultOption.value = '';
                defaultOption.textContent = '全部 ' + header;
                select.appendChild(defaultOption);
                
                // 获取该列的唯一值并添加到下拉框，但限制数量以提高性能
                try {
                    if (data && ExcelProcessor && typeof ExcelProcessor.getColumnUniqueValues === 'function') {
                        const uniqueValues = ExcelProcessor.getColumnUniqueValues(data, header);
                        if (uniqueValues && Array.isArray(uniqueValues)) {
                            // 限制下拉框选项数量以提高性能
                            const maxOptions = 100;
                            const displayValues = uniqueValues.slice(0, maxOptions);
                            
                            displayValues.forEach(value => {
                                const option = document.createElement('option');
                                option.value = value;
                                option.textContent = value;
                                select.appendChild(option);
                            });
                            
                            // 如果有更多选项，添加提示
                            if (uniqueValues.length > maxOptions) {
                                const moreOption = document.createElement('option');
                                moreOption.disabled = true;
                                moreOption.textContent = '... 还有 ' + (uniqueValues.length - maxOptions) + ' 个选项';
                                select.appendChild(moreOption);
                            }
                        }
                    }
                } catch (error) {
                    console.warn('获取列 ' + header + ' 的唯一值时出错:', error);
                }
                
                // 添加"自定义输入"选项
                const customOption = document.createElement('option');
                customOption.value = '__custom__';
                customOption.textContent = '自定义输入...';
                select.appendChild(customOption);
                
                // 添加输入框（默认隐藏）
                const input = document.createElement('input');
                input.type = 'text';
                input.placeholder = '输入筛选条件';
                input.className = 'custom-input hidden';
                
                // 监听选择变化
                select.addEventListener('change', (e) => {
                    if (e.target.value === '__custom__') {
                        input.classList.remove('hidden');
                        input.focus();
                        filters[header] = '';
                    } else {
                        input.classList.add('hidden');
                        filters[header] = e.target.value;
                    }
                    
                    if (onFilterChange) {
                        onFilterChange(filters);
                    }
                });
                
                // 监听输入变化
                input.addEventListener('input', (e) => {
                    filters[header] = e.target.value;
                    if (onFilterChange) {
                        onFilterChange(filters);
                    }
                });
                
                filterGroup.appendChild(label);
                filterGroup.appendChild(select);
                filterGroup.appendChild(input);
                filterContainer.appendChild(filterGroup);
            });
            
            container.appendChild(filterContainer);
        } catch (error) {
            console.error('创建筛选控件时出错:', error);
            container.innerHTML = '<p class="error">创建筛选控件时出错: ' + error.message + '</p>';
        }
    }
    
    /**
     * 更新筛选控件的值
     * @param {Object} filters - 当前筛选条件
     * @param {HTMLElement} container - 筛选控件容器
     */
    static updateFilterControls(filters, container) {
        if (!filters || !container) {
            return;
        }
        
        try {
            Object.keys(filters).forEach(key => {
                const select = container.querySelector('#filter-' + key);
                if (select) {
                    select.value = filters[key] || '';
                }
            });
        } catch (error) {
            console.error('更新筛选控件时出错:', error);
        }
    }
    
    /**
     * 应用语法高亮到代码块
     * @param {HTMLElement} container - 包含代码块的容器元素
     */
    static applySyntaxHighlighting(container) {
        // 确保Prism.js已加载
        if (typeof Prism !== 'undefined') {
            // 查找所有带有language-java类的代码块
            const codeBlocks = container.querySelectorAll('pre.language-java');
            codeBlocks.forEach(block => {
                // 应用语法高亮
                Prism.highlightElement(block);
            });
        }
    }
}