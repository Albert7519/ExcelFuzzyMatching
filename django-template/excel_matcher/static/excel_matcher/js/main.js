document.addEventListener('DOMContentLoaded', function() {
    // DOM元素
    const uploadForm = document.getElementById('upload-form');
    const fileUpload = document.getElementById('file-upload');
    const uploadBtn = document.getElementById('upload-btn');
    const uploadStatus = document.getElementById('upload-status');
    const uploadLoader = document.getElementById('upload-loader');
    const columnsSection = document.getElementById('columns-section');
    const columnsList = document.getElementById('columns-list');
    const thresholdInput = document.getElementById('threshold');
    const thresholdValue = document.getElementById('threshold-value');
    const processBtn = document.getElementById('process-btn');
    const processLoader = document.getElementById('process-loader');
    const processStatus = document.getElementById('process-status');
    const downloadSection = document.getElementById('download-section');
    const downloadBtn = document.getElementById('download-btn');
    
    // 事件监听
    uploadForm.addEventListener('submit', handleFileUpload);
    processBtn.addEventListener('click', processFile);
    thresholdInput.addEventListener('input', updateThresholdValue);
    
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
                showStatus(uploadStatus, '文件上传成功！请选择需要模糊匹配的列。', 'success');
                
                // 显示列选择区域
                columnsSection.classList.remove('hidden');
                
                // 显示列选择
                displayColumns(data.columns);
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
    
    // 显示列选择
    function displayColumns(columns) {
        columnsList.innerHTML = '';
        
        columns.forEach(column => {
            const label = document.createElement('label');
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.name = 'column';
            checkbox.value = column;
            
            label.appendChild(checkbox);
            label.appendChild(document.createTextNode(' ' + column));
            
            columnsList.appendChild(label);
        });
    }
    
    // 处理文件
    function processFile() {
        // 获取选中的列
        const selectedColumns = Array.from(document.querySelectorAll('input[name="column"]:checked'))
            .map(checkbox => checkbox.value);
        
        // 获取阈值
        const threshold = parseInt(thresholdInput.value);
        
        if (selectedColumns.length === 0) {
            showStatus(processStatus, '请至少选择一列进行处理！', 'error');
            return;
        }
        
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
                threshold: threshold
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