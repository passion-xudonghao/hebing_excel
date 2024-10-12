// app.js

let excelData = [];

// 监听文件上传
document.getElementById('fileInput').addEventListener('change', handleFileUpload);

function handleFileUpload(event) {
    const files = event.target.files;
    if (files.length === 0) {
        alert('请上传文件');
        return;
    }

    // 遍历上传的文件
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // 获取第一个工作表
            const sheetName = workbook.SheetNames[0];
            const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

            // 将表格内容转换为JSON数组
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

            // 为每个文件数据新增一个字段“数据类型”，设置为文件名
            jsonData.forEach(row => {
                row['品牌分类'] = file.name;
				
				// 从订单日期字段中提取年月日并新增为一个新的"日期"字段
				    if (row['添加日期']) {
				        const dateValue = new Date(row['添加日期']); // 将订单日期转为 Date 对象
				
				        // 格式化为 YYYY-MM-DD 格式
				        const year = dateValue.getFullYear();
				        const month = (dateValue.getMonth() + 1).toString().padStart(2, '0'); // 月份从 0 开始，所以要加 1，并确保两位数格式
				        const day = dateValue.getDate().toString().padStart(2, '0'); // 确保日期两位数格式
				
				        const formattedDate = `${year}-${month}-${day}`; // 组合成 YYYY-MM-DD 格式
				
				        // 添加新的字段"日期"
				        row['日期'] = formattedDate;
				    } else {
				        // 如果没有时间信息，可以设置为空或者默认值
				        row['日期'] = '无时间信息';
				    }
            });

            // 将处理后的数据存储到全局数组
            excelData.push(...jsonData);

            // 显示每个文件的内容
            document.getElementById('output').innerHTML += `<p>${file.name} 已上传并处理。</p>`;
        };
        reader.readAsArrayBuffer(file);
    }
}

// 合并文件并导出为新的Excel
function mergeExcelFiles() {
    if (excelData.length === 0) {
        alert('没有数据可合并');
        return;
    }

    // 将合并后的数据转换为工作表
    const worksheet = XLSX.utils.json_to_sheet(excelData);

    // 创建一个新的工作簿
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, worksheet, 'ALL_TABLE');

    // 导出Excel文件
    XLSX.writeFile(newWorkbook, 'ALL_TABLE.xlsx');
}
