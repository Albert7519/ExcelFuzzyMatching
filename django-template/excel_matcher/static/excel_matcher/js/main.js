document.addEventListener('DOMContentLoaded', function() {
    // DOM元素 - 基本界面
    const uploadForm = document.getElementById('upload-form');
    const fileUpload = document.getElementById('file-upload');
    const uploadBtn = document.getElementById('upload-btn');
    const uploadStatus = document.getElementById('upload-status');
    const uploadLoader = document.getElementById('upload-loader');
    
    // DOM元素 - 模式选择
    const modeSection = document.getElementById('mode-section');
    const modeContinueBtn = document.getElementById('mode-continue-btn');
    const modeRadios = document.querySelectorAll('input[name="processing-mode"]');
    
    // DOM元素 - 列选择
    const columnsSection = document.getElementById('columns-section');
    const selfLearningColumns = document.getElementById('self-learning-columns');
    const referenceColumns = document.getElementById('reference-columns');
    const columnsList = document.getElementById('columns-list');
    const referenceColumnSelect = document.getElementById('reference-column-select');
    const matchColumnsList = document.getElementById('match-columns-list');
    
    // DOM元素 - 阈值设置和处理
    const thresholdInput = document.getElementById('threshold');
    const thresholdValue = document.getElementById('threshold-value');
    const processBtn = document.getElementById('process-btn');
    const processLoader = document.getElementById('process-loader');
    const processStatus = document.getElementById('process-status');
    
    // DOM元素 - 下载
    const downloadSection = document.getElementById('download-section');
    const downloadBtn = document.getElementById('download-btn');
    
    // 状态变量
    let currentMode = 'SELF_LEARNING'; // 默认模式
    let availableColumns = []; // 可用的列
    
    // 事件监听
    uploadForm.addEventListener('submit', handleFileUpload);
    modeContinueBtn.addEventListener('click', handleModeContinue);
    processBtn.addEventListener('click', processFile);
    thresholdInput.addEventListener('input', updateThresholdValue);
    
    // 模式选择事件监听
    modeRadios.forEach(radio => {
        radio.addEventListener('change', function() {
            currentMode = this.value;
        });
    });
    
    // 标准列选择事件监听
    referenceColumnSelect.addEventListener('change', updateMatchColumnsList);
    
    // 更新阈值显示
    function updateThresholdValue() {
        thresholdValue.textContent = thresholdInput.value;
    }
    
    // 处理文件上传
    function handleFileUpload(event) {
        event.preventDefault(); // 阻止表单默认提交
        
        const file = fileUpload.files[0];
        if (!file) {
            showStatus(uploadStatus, '请选择一个Excel文件', 'error');
            return;
        }
        
        // 检查文件类型
        if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
            showStatus(uploadStatus, '请上传Excel文件(.xlsx或.xls)', 'error');
            return;
        }
        
        // 显示加载状态
        uploadLoader.classList.remove('hidden');
        uploadStatus.classList.add('hidden');
        
        const formData = new FormData(uploadForm);
        
        fetch('/excel/upload/', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            uploadLoader.classList.add('hidden');
            
            if (data.success) {
                // 显示成功消息
                showStatus(uploadStatus, '文件上传成功！请选择处理模式。', 'success');
                
                // 保存可用列
                availableColumns = data.columns;
                
                // 显示模式选择区域
                modeSection.classList.remove('hidden');
            } else {
                // 显示错误消息
                showStatus(uploadStatus, '上传失败：' + data.error, 'error');
            }
        })
        .catch(error => {
            uploadLoader.classList.add('hidden');
            showStatus(uploadStatus, '上传出错：' + error.message, 'error');
        });
    }
    
    // 处理模式选择继续按钮
    function handleModeContinue() {
        // 根据当前模式显示相应的列选择界面
        if (currentMode === 'SELF_LEARNING') {
            selfLearningColumns.classList.remove('hidden');
            referenceColumns.classList.add('hidden');
            
            // 显示所有列供选择
            displayColumns(columnsList, availableColumns);
        } else {
            selfLearningColumns.classList.add('hidden');
            referenceColumns.classList.remove('hidden');
            
            // 填充标准列下拉选择框
            populateReferenceSelect(availableColumns);
            
            // 初始化匹配列列表为空
            matchColumnsList.innerHTML = '';
        }
        
        // 显示列选择区域
        columnsSection.classList.remove('hidden');
    }
    
    // 更新标准列选择后的匹配列列表
    function updateMatchColumnsList() {
        const referenceColumn = referenceColumnSelect.value;
        if (!referenceColumn) {
            matchColumnsList.innerHTML = '';
            return;
        }
        
        // 过滤掉标准列，其余都可以作为匹配列
        const matchColumns = availableColumns.filter(col => col !== referenceColumn);
        displayColumns(matchColumnsList, matchColumns);
    }
    
    // 填充标准列下拉选择框
    function populateReferenceSelect(columns) {
        // 清空现有选项，只保留默认选项
        referenceColumnSelect.innerHTML = '<option value="">-- 请选择 --</option>';
        
        // 添加所有列作为选项
        columns.forEach(column => {
            const option = document.createElement('option');
            option.value = column;
            option.textContent = column;
            referenceColumnSelect.appendChild(option);
        });
    }
    
    // 显示列选择
    function displayColumns(container, columns) {
        container.innerHTML = '';
        
        columns.forEach(column => {
            const label = document.createElement('label');
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.name = 'column';
            checkbox.value = column;
            
            label.appendChild(checkbox);
            label.appendChild(document.createTextNode(' ' + column));
            
            container.appendChild(label);
        });
    }
    
    // 处理文件
    function processFile() {
        let selectedColumns = [];
        let referenceColumn = null;
        
        // 根据当前模式获取选中的列
        if (currentMode === 'SELF_LEARNING') {
            selectedColumns = Array.from(
                document.querySelectorAll('#columns-list input[name="column"]:checked')
            ).map(checkbox => checkbox.value);
            
            if (selectedColumns.length === 0) {
                showStatus(processStatus, '请至少选择一列进行处理！', 'error');
                return;
            }
        } else {
            // 参照标准匹配模式
            referenceColumn = referenceColumnSelect.value;
            
            if (!referenceColumn) {
                showStatus(processStatus, '请选择一个标准参照列！', 'error');
                return;
            }
            
            // 获取选中的匹配列
            const matchColumns = Array.from(
                document.querySelectorAll('#match-columns-list input[name="column"]:checked')
            ).map(checkbox => checkbox.value);
            
            if (matchColumns.length === 0) {
                showStatus(processStatus, '请至少选择一列进行匹配！', 'error');
                return;
            }
            
            // 标准列也需要加入处理列表
            selectedColumns = [referenceColumn, ...matchColumns];
        }
        
        // 获取阈值
        const threshold = parseInt(thresholdInput.value);
        
        // 显示加载状态
        processLoader.classList.remove('hidden');
        processStatus.classList.add('hidden');
        
        fetch('/excel/process/', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'X-CSRFToken': getCookie('csrftoken')
            },
            body: JSON.stringify({
                columns_to_match: selectedColumns,
                threshold: threshold,
                processing_mode: currentMode,
                reference_column: referenceColumn
            })
        })
        .then(response => response.json())
        .then(data => {
            processLoader.classList.add('hidden');
            
            if (data.success) {
                // 显示成功消息
                showStatus(processStatus, '文件处理成功！请点击下载按钮获取处理后的文件。', 'success');
                
                // 显示下载区域
                downloadSection.classList.remove('hidden');
            } else {
                // 显示错误消息
                showStatus(processStatus, '处理失败：' + data.error, 'error');
            }
        })
        .catch(error => {
            processLoader.classList.add('hidden');
            showStatus(processStatus, '处理出错：' + error.message, 'error');
        });
    }
    
    // 显示状态消息
    function showStatus(element, message, type) {
        element.textContent = message;
        element.className = 'status ' + type;
        element.classList.remove('hidden');
    }
    
    // 获取CSRF Token
    function getCookie(name) {
        let cookieValue = null;
        if (document.cookie && document.cookie !== '') {
            const cookies = document.cookie.split(';');
            for (let i = 0; i < cookies.length; i++) {
                const cookie = cookies[i].trim();
                if (cookie.substring(0, name.length + 1) === (name + '=')) {
                    cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
                    break;
                }
            }
        }
        return cookieValue;
    }
});