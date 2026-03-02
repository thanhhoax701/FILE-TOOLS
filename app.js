// Read Excel file
async function readExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

            // if the sheet contains a fullname column, always ensure there is a
            // CUSTOMER_NAME column with the same value. this fixes situations
            // where the source file only provides "FullName" but later
            // processing expects CUSTOMER_NAME.
            if (json.length > 0) {
                const headers = Object.keys(json[0]);
                const fullNameKey = headers.find(h => h.toLowerCase().replace(/[_\s]/g, '') === 'fullname');
                if (fullNameKey) {
                    json.forEach(row => {
                        row['CUSTOMER_NAME'] = row[fullNameKey];
                    });
                }

                // ensure ServiceType column exists with default 'KTDK'
                // if the user provides their own value it should not be overwritten
                json.forEach(row => {
                    if (row['ServiceType'] == null || String(row['ServiceType']).trim() === '') {
                        row['ServiceType'] = 'KTDK';
                    }
                });
            }

            resolve(json);
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// helper to switch between UI sections
function showSection(section) {
    document.getElementById('compareSection').style.display = section === 'compare' ? 'block' : 'none';
    document.getElementById('splitSection').style.display = section === 'split' ? 'block' : 'none';
    document.getElementById('splitDateSection').style.display = section === 'splitDate' ? 'block' : 'none';
    document.getElementById('navCompare').classList.toggle('active', section === 'compare');
    document.getElementById('navSplit').classList.toggle('active', section === 'split');
    document.getElementById('navSplitDate').classList.toggle('active', section === 'splitDate');
}

// Normalize value for comparison
function normalize(value) {
    return String(value).trim().toUpperCase();
}

// Get unique values from a column
function getUniqueValues(data, columnName) {
    const uniqueSet = new Set();
    const uniqueList = [];
    data.forEach(row => {
        const value = normalize(row[columnName]);
        if (value && !uniqueSet.has(value)) {
            uniqueSet.add(value);
            uniqueList.push({
                value: value,
                display: row[columnName],
                fullRow: row
            });
        }
    });
    return { set: uniqueSet, list: uniqueList };
}

// Display table
function displayTable(containerId, data, columns, highlightValues = null) {
    const container = document.getElementById(containerId);
    container.innerHTML = '';

    if (!data || data.length === 0) {
        container.innerHTML = '<p style="color:#999; text-align:center;">Không có dữ liệu</p>';
        return;
    }

    // Detect date columns by name pattern
    const dateColumns = new Set();
    columns.forEach(col => {
        if (col.toLowerCase().includes('date') || col.toLowerCase().includes('remind')) {
            dateColumns.add(col);
        }
    });

    let html = '<table style="width:100%; border-collapse:collapse; font-size:12px;">';

    // Header
    html += '<tr style="background:#f0f0f0; border-bottom:2px solid #ccc;">';
    columns.forEach(col => {
        html += `<th style="padding:8px; border:1px solid #ddd; text-align:left;">${col}</th>`;
    });
    html += '</tr>';

    // Rows
    data.forEach((row, idx) => {
        let rowStyle = '';
        if (row._diff) {
            rowStyle = 'background:#ffcccc; font-weight:bold;'; // Highlight different rows
        }

        html += `<tr style="border-bottom:1px solid #eee; ${rowStyle}">`;
        columns.forEach(col => {
            let value = row[col] || '';

            // Format date columns
            if (dateColumns.has(col)) {
                value = formatDateValue(value);
            }

            html += `<td style="padding:8px; border:1px solid #ddd; word-break:break-word;">${value}</td>`;
        });
        html += '</tr>';
    });

    html += '</table>';
    container.innerHTML = html;
}

// Preview compare files when they are selected (show contents without diff)
async function previewCompareFiles() {
    const file1 = document.getElementById("compareFile1").files[0];
    const file2 = document.getElementById("compareFile2").files[0];
    const resultDiv = document.getElementById("result");

    let data1 = [];
    let data2 = [];

    if (file1) {
        try {
            data1 = await readExcel(file1);
            const cols1 = data1.length ? Object.keys(data1[0]) : [];
            displayTable("table1", data1, cols1);
            document.getElementById("statsFile1").innerText = `File 1: ${data1.length} dòng`;
        } catch (e) {
            console.error(e);
            alert("Không thể đọc File 1!");
        }
    }

    if (file2) {
        try {
            data2 = await readExcel(file2);
            const cols2 = data2.length ? Object.keys(data2[0]) : [];
            displayTable("table2", data2, cols2);
            document.getElementById("statsFile2").innerText = `File 2: ${data2.length} dòng`;
        } catch (e) {
            console.error(e);
            alert("Không thể đọc File 2!");
        }
    }

    // Hide diff since this is only preview
    document.getElementById("diffSection").style.display = 'none';
    document.getElementById("exportBtn").style.display = 'none';
    resultDiv.style.display = (data1.length || data2.length) ? 'block' : 'none';
}

// Main process
async function processFiles() {
    const file1 = document.getElementById("compareFile1").files[0];
    const file2 = document.getElementById("compareFile2").files[0];
    const columnName = document.getElementById("columnName").value.trim();
    const resultDiv = document.getElementById("result");

    if (!file1 || !file2 || !columnName) {
        alert("Vui lòng chọn đủ 2 file và nhập tên cột để so sánh!");
        return;
    }

    resultDiv.style.display = 'none';
    document.getElementById("statusCompare").innerText = "Đang xử lý...";

    try {
        const data1 = await readExcel(file1);
        const data2 = await readExcel(file2);

        if (data1.length === 0 || data2.length === 0) {
            alert("Một trong hai file không có dữ liệu!");
            document.getElementById("statusCompare").innerText = "";
            return;
        }

        if (!(columnName in data1[0]) || !(columnName in data2[0])) {
            alert("Không tìm thấy cột '" + columnName + "' trong một trong hai file!");
            document.getElementById("statusCompare").innerText = "";
            return;
        }

        // Get columns from both files
        const columns1 = Object.keys(data1[0]);
        const columns2 = Object.keys(data2[0]);

        // Get unique values from each file
        const file1Values = getUniqueValues(data1, columnName);
        const file2Values = getUniqueValues(data2, columnName);

        // Find differences: rows in file2 that are not in file1
        const diffData = [];
        file2Values.list.forEach(item => {
            if (!file1Values.set.has(item.value)) {
                const rowData = item.fullRow;
                rowData._diff = true; // Mark as difference
                diffData.push(rowData);
            }
        });

        // Update statistics
        document.getElementById("statsFile1").innerText = `File 1: ${data1.length} dòng`;
        document.getElementById("statsFile2").innerText = `File 2: ${data2.length} dòng`;
        document.getElementById("statsDiff").innerText = `Lệch: ${diffData.length} dòng trong File 2 không có trong File 1`;

        // Display tables
        // Always show the full contents of both files
        displayTable("table1", data1, columns1);
        displayTable("table2", data2, columns2);

        // Display difference details and toggle export button
        if (diffData.length > 0) {
            displayTable("tableDiff", diffData, columns2);
            document.getElementById("diffSection").style.display = 'block';
            document.getElementById("exportBtn").style.display = 'inline-block';
        } else {
            document.getElementById("diffSection").style.display = 'none';
            document.getElementById("exportBtn").style.display = 'none';
        }

        resultDiv.style.display = 'block';
        document.getElementById("status").innerText = "";

    } catch (err) {
        console.error(err);
        alert("Có lỗi khi xử lý file!");
        document.getElementById("status").innerText = "";
    }
}

// Export difference to Excel
function exportDifference() {
    const table2Content = document.getElementById("table2").innerHTML;
    if (!table2Content || table2Content.includes("Không có dữ liệu")) {
        alert("Không có dữ liệu để xuất!");
        return;
    }

    const file2 = document.getElementById("compareFile2").files[0];
    readExcel(file2).then(data2 => {
        const columnName = document.getElementById("columnName").value.trim();
        const file1 = document.getElementById("compareFile1").files[0];

        readExcel(file1).then(data1 => {
            const file1Values = getUniqueValues(data1, columnName);

            const diffData = [];
            data2.forEach(row => {
                const value = normalize(row[columnName]);
                if (value && !file1Values.set.has(value)) {
                    diffData.push(row);
                }
            });

            if (diffData.length > 0) {
                const ws = XLSX.utils.json_to_sheet(diffData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "DuLieuLech");
                XLSX.writeFile(wb, "du_lieu_lech_file2.xlsx");
            }
        });
    });
}

// Split a file into multiple sheets, each with up to `rowsPerSheet` rows
async function splitFileIntoSheets(file, rowsPerSheet = 1000) {
    if (!file) {
        alert('Vui lòng chọn file để tách!');
        return;
    }

    try {
        document.getElementById('statusSplit').innerText = 'Đang đọc và tách file...';
        const data = await readExcel(file);

        if (!data || data.length === 0) {
            alert('File không có dữ liệu để tách.');
            document.getElementById('statusSplit').innerText = '';
            return;
        }

        // Display a preview (first 100 rows) so user sees the file
        const cols = Object.keys(data[0]);
        const previewRows = data.slice(0, 100);
        displayTable('tableSplit', previewRows, cols);
        // preview only; no stats element

        // Create workbook and append chunked sheets
        const wb = XLSX.utils.book_new();
        let sheetCount = 0;
        for (let i = 0; i < data.length; i += rowsPerSheet) {
            const chunk = data.slice(i, i + rowsPerSheet);
            const ws = XLSX.utils.json_to_sheet(chunk);
            const start = i + 1;
            const end = Math.min(i + rowsPerSheet, data.length);
            // include number of rows in the sheet name
            const sheetName = `${start}-${end} (${chunk.length})`;
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
            sheetCount++;
        }

        const outName = file.name.replace(/\.[^.]+$/, '') + `_split_${rowsPerSheet}rows.xlsx`;
        XLSX.writeFile(wb, outName);

        document.getElementById('statusSplit').innerText = `Hoàn thành: tách thành ${sheetCount} sheet. Tệp đã tải về: ${outName}`;
    } catch (err) {
        console.error(err);
        alert('Có lỗi khi tách file: ' + err.message);
        document.getElementById('statusSplit').innerText = '';
    }
}

// read and preview file in split section without splitting
async function readSplitFile() {
    const file = document.getElementById('splitFile').files[0];
    if (!file) {
        alert('Vui lòng chọn file để đọc!');
        return;
    }
    try {
        document.getElementById('statusSplit').innerText = 'Đang đọc file...';
        const data = await readExcel(file);
        if (!data || data.length === 0) {
            alert('File không có dữ liệu.');
            document.getElementById('statusSplit').innerText = '';
            return;
        }
        const cols = Object.keys(data[0]);
        displayTable('tableSplit', data.slice(0, 100), cols); // show first 100 rows
        document.getElementById('statusSplit').innerText = `Đã đọc ${data.length} dòng`;
    } catch (err) {
        console.error(err);
        alert('Lỗi khi đọc file: ' + err.message);
        document.getElementById('statusSplit').innerText = '';
    }
}

// Convert Excel serial date to readable date string
function formatDateValue(dateVal) {
    // If it's a number (Excel serial), convert it
    if (typeof dateVal === 'number') {
        // Excel date serial starts from 1900-01-01
        const excelEpoch = new Date(1900, 0, 1);
        const date = new Date(excelEpoch.getTime() + (dateVal - 1) * 24 * 60 * 60 * 1000);
        // Return in YYYY-MM-DD format
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }
    // If already a string, return as-is
    return String(dateVal).trim();
}

// Load unique dates from selected file
async function loadDatesList() {
    const file = document.getElementById('splitDateFile').files[0];
    const dateColName = document.getElementById('dateColumn').value.trim();

    if (!file) {
        alert('Vui lòng chọn file!');
        return;
    }
    if (!dateColName) {
        alert('Vui lòng nhập tên cột ngày!');
        return;
    }

    try {
        document.getElementById('statusSplitDate').innerText = 'Đang đọc ngày...';
        const data = await readExcel(file);

        if (!data || data.length === 0) {
            alert('File không có dữ liệu.');
            document.getElementById('statusSplitDate').innerText = '';
            return;
        }

        // Check if date column exists
        if (!(dateColName in data[0])) {
            alert(`Không tìm thấy cột "${dateColName}" trong file!`);
            document.getElementById('statusSplitDate').innerText = '';
            return;
        }

        // Store data globally for export
        window.splitDateData = data;

        // Get unique dates with proper formatting
        const uniqueDates = new Set();
        data.forEach(row => {
            const dateVal = row[dateColName];
            if (dateVal) {
                const formatted = formatDateValue(dateVal);
                uniqueDates.add(formatted);
            }
        });

        const sortedDates = Array.from(uniqueDates).sort();

        // Display dates with checkboxes
        let html = '';
        sortedDates.forEach((date, idx) => {
            const rows = data.filter(r => formatDateValue(r[dateColName]) === date).length;
            html += `<label style="display:inline-block; margin:10px 20px 10px 0; cursor:pointer;">
                <input type="checkbox" class="dateCheck" value="${date}">
                ${date} (${rows} dòng)
            </label>`;
        });

        document.getElementById('datesList').innerHTML = html;
        document.getElementById('datesContainer').style.display = 'block';
        document.getElementById('exportDateBtn').style.display = 'inline-block';
        document.getElementById('statusSplitDate').innerText = `Tìm thấy ${sortedDates.length} ngày khác nhau`;

        // Store date column name globally
        window.dateColumnName = dateColName;

    } catch (err) {
        console.error(err);
        alert('Lỗi: ' + err.message);
        document.getElementById('statusSplitDate').innerText = '';
    }
}

// Export selected dates to workbook (one sheet per date)
function exportByDate() {
    if (!window.splitDateData || !window.dateColumnName) {
        alert('Vui lòng đọc file trước!');
        return;
    }

    // Get selected dates from checkboxes
    const checkboxes = document.querySelectorAll('.dateCheck:checked');
    if (checkboxes.length === 0) {
        alert('Vui lòng chọn ít nhất một ngày!');
        return;
    }

    try {
        document.getElementById('statusSplitDate').innerText = 'Đang xuất...';
        const wb = XLSX.utils.book_new();
        const dateColName = window.dateColumnName;
        const data = window.splitDateData;

        // Process each selected date
        checkboxes.forEach((checkbox) => {
            const selectedDate = checkbox.value;
            const filteredRows = data.filter(row =>
                formatDateValue(row[dateColName]) === selectedDate
            );

            if (filteredRows.length > 0) {
                // ensure service type default for each row
                filteredRows.forEach(r => {
                    if (r['ServiceType'] == null || String(r['ServiceType']).trim() === '') {
                        r['ServiceType'] = 'KTDK';
                    }
                });

                const ws = XLSX.utils.json_to_sheet(filteredRows);
                XLSX.utils.book_append_sheet(wb, ws, selectedDate);
            }
        });

        const fileName = `KTDK_${new Date().toISOString().split('T')[0]}.xlsx`;
        XLSX.writeFile(wb, fileName);

        document.getElementById('statusSplitDate').innerText = `Đã xuất ${checkboxes.length} ngày vào: ${fileName}`;
    } catch (err) {
        console.error(err);
        alert('Lỗi khi xuất: ' + err.message);
        document.getElementById('statusSplitDate').innerText = '';
    }
}
function splitFile2() {
    const file = document.getElementById('splitFile').files[0];
    const rows = parseInt(document.getElementById('rowsPerSheet').value, 10) || 1000;
    splitFileIntoSheets(file, rows);
}

// ========== KTĐK Split Functionality ==========

// Step 1: Read KTĐK file and display all data
async function readKTDKFile() {
    const file = document.getElementById('splitDateFile').files[0];

    if (!file) {
        alert('Vui lòng chọn file KTĐK!');
        return;
    }

    try {
        document.getElementById('statusKTDKRead').innerText = 'Đang đọc file...';

        // Convert to JSON for data processing
        const data = await readExcel(file);

        // Create 2 standard header rows according to template format
        const headerRow1 = ['FullName', 'PhoneNo', 'ServiceType', 'ServiceDate', 'ServiceName', 'ServiceCode', 'BranchCode', 'Dist', 'Province', 'FName_1', 'FName_2', 'FName_3', 'FName_4', 'FName_5', 'FName_6', 'FName_7', 'FName_8', 'FName_9', 'FName_10'];
        const headerRow2 = ['FullName', 'PhoneNo', 'ServiceType', 'ServiceDate', 'ServiceName', 'ServiceCode', 'BranchCode', 'Dist', 'Province', 'MODEL_NAME', 'TEST_TYPE', 'CONTENT', 'ENDDATE', 'PLACENAME', 'NUMBER_PLATE', 'CUSTOMER_NAME', 'FName_8', 'FName_9', 'FName_10'];
        window.ktdkHeaderRows = [headerRow1, headerRow2];

        if (!data || data.length === 0) {
            alert('File không có dữ liệu.');
            document.getElementById('statusKTDKRead').innerText = '';
            return;
        }

        // Store data globally
        window.ktdkData = data;
        // clear any previous split results
        window.ktdkSplitResults = null;
        window.ktdkSplitResultsB = null;
        document.getElementById('ktdkTypeA_DatesContainer').style.display = 'none';
        document.getElementById('ktdkTypeA_ResultsContainer').style.display = 'none';
        document.getElementById('ktdkTypeB_DatesContainer').style.display = 'none';
        document.getElementById('ktdkTypeB_ResultsContainer').style.display = 'none';

        // Display all data
        const cols = Object.keys(data[0]);
        displayTable('tableKTDKFull', data, cols);

        // Show next step UI
        document.getElementById('ktdkDataPreview').style.display = 'block';
        document.getElementById('ktdkTypeA_SelectionStep').style.display = 'block';
        document.getElementById('ktdkTypeB_SelectionStep').style.display = 'block';
        document.getElementById('kdkTypeB_SettingsAfterTemplate').style.display = 'none'; // Hide settings until template is loaded
        document.getElementById('statusKTDKRead').innerText = `Đã tải ${data.length} dòng`;

    } catch (err) {
        console.error(err);
        alert('Lỗi khi đọc file: ' + err.message);
        document.getElementById('statusKTDKRead').innerText = '';
    }
}

// Step 2: Load and display dates
async function selectDatesToSplitKTDK() {
    const dateColName = document.getElementById('kdkDateColumn').value.trim();

    if (!window.ktdkData) {
        alert('Vui lòng đọc file trước!');
        return;
    }

    if (!dateColName) {
        alert('Vui lòng nhập tên cột ngày!');
        return;
    }

    const data = window.ktdkData;

    // Check column exists
    if (!(dateColName in data[0])) {
        alert(`Không tìm thấy cột "${dateColName}" trong file!`);
        return;
    }

    try {
        document.getElementById('statusKTDKSelect').innerText = 'Đang tải danh sách ngày...';

        // Get unique dates with proper formatting
        const uniqueDates = new Set();
        data.forEach(row => {
            const dateVal = row[dateColName];
            if (dateVal) {
                const formatted = formatDateValue(dateVal);
                uniqueDates.add(formatted);
            }
        });

        const sortedDates = Array.from(uniqueDates).sort();

        // Display dates with checkboxes
        let html = '';
        sortedDates.forEach((date) => {
            const rows = data.filter(r => formatDateValue(r[dateColName]) === date).length;
            html += `<label style="display:inline-block; margin:10px 20px 10px 0; cursor:pointer;">
                <input type="checkbox" class="kdkDateCheck" value="${date}">
                ${date} (${rows} dòng)
            </label>`;
        });

        document.getElementById('kdkDatesList').innerHTML = html;
        // add select-all handler
        document.getElementById('kdkDateSelectAll').addEventListener('change', function () {
            const checked = this.checked;
            document.querySelectorAll('.kdkDateCheck').forEach(cb => cb.checked = checked);
        });

        document.getElementById('ktdkTypeA_DatesContainer').style.display = 'block';
        document.getElementById('statusKTDKSelect').innerText = `Tìm thấy ${sortedDates.length} ngày khác nhau`;

        // Store column name globally
        window.kdkDateColumnName = dateColName;

    } catch (err) {
        console.error(err);
        alert('Lỗi: ' + err.message);
        document.getElementById('statusKTDKSelect').innerText = '';
    }
}

// Step 3: Split by selected dates into 2 consolidated workbooks
async function doSplitKTDK() {
    if (!window.ktdkData || !window.kdkDateColumnName) {
        alert('Vui lòng đọc file và chọn ngày trước!');
        return;
    }

    const checkboxes = document.querySelectorAll('.kdkDateCheck:checked');
    if (checkboxes.length === 0) {
        alert('Vui lòng chọn ít nhất một ngày!');
        return;
    }

    try {
        document.getElementById('statusKTDKSelect').innerText = 'Đang tạo file tổng hợp...';

        // Helper function to find column name case-insensitively
        function findColumnName(data, searchNames) {
            if (!data || data.length === 0) return null;
            const headers = Object.keys(data[0]);
            for (let searchName of searchNames) {
                const found = headers.find(h => h.toLowerCase() === searchName.toLowerCase());
                if (found) return found;
            }
            return null;
        }

        const data = window.ktdkData;
        const dateColName = window.kdkDateColumnName;

        // Find actual column names in the data
        const actualSet1Cols = [];
        const actualSet2Cols = [];

        // For set1 (CONTACT)
        if (findColumnName(data, ['FULLNAME', 'FullName', 'Full_Name', 'CUSTOMER_NAME', 'CustomerName', 'Customer_Name']))
            actualSet1Cols.push(findColumnName(data, ['FULLNAME', 'FullName', 'Full_Name', 'CUSTOMER_NAME', 'CustomerName', 'Customer_Name']));
        if (findColumnName(data, ['TELEINFO', 'TeleInfo', 'PHONE', 'MOBILE']))
            actualSet1Cols.push(findColumnName(data, ['TELEINFO', 'TeleInfo', 'PHONE', 'MOBILE']));
        if (findColumnName(data, ['BRANCHCODE', 'BranchCode', 'Branch_Code', 'BRANCH_CODE', 'Branch']))
            actualSet1Cols.push(findColumnName(data, ['BRANCHCODE', 'BranchCode', 'Branch_Code', 'BRANCH_CODE', 'Branch']));

        // For set2 (TEST)
        if (findColumnName(data, ['MODEL_NAME', 'ModelName']))
            actualSet2Cols.push(findColumnName(data, ['MODEL_NAME', 'ModelName']));
        if (findColumnName(data, ['TEST_TYPE', 'TestType']))
            actualSet2Cols.push(findColumnName(data, ['TEST_TYPE', 'TestType']));
        if (findColumnName(data, ['CONTENT', 'Content']))
            actualSet2Cols.push(findColumnName(data, ['CONTENT', 'Content']));
        if (findColumnName(data, ['ENDDATE', 'END_DATE', 'EndDate']))
            actualSet2Cols.push(findColumnName(data, ['ENDDATE', 'END_DATE', 'EndDate']));
        if (findColumnName(data, ['PLACE', 'Place']))
            actualSet2Cols.push(findColumnName(data, ['PLACE', 'Place']));
        if (findColumnName(data, ['NUMBER_PLATE', 'NumberPlate', 'Number_Plate']))
            actualSet2Cols.push(findColumnName(data, ['NUMBER_PLATE', 'NumberPlate', 'Number_Plate']));

        // Prepare results by date for each column set
        window.ktdkSplitResults = {
            set1: {
                cols: actualSet1Cols,
                dateMap: {} // key = displayDate, value = array of rows for CONTACT
            },
            set2: {
                cols: actualSet2Cols,
                dateMap: {} // key = displayDate, value = array of rows for TEST
            }
        };
        let processedCount = 0;

        // Process each selected date and collect data into maps
        checkboxes.forEach((checkbox) => {
            const selectedDate = checkbox.value;
            const displayDate = selectedDate.replace(/(\d{4})-(\d{2})-(\d{2})/, '$3-$2-$1');

            // Filter data for this date
            const filteredRows = data.filter(row =>
                formatDateValue(row[dateColName]) === selectedDate
            );

            if (filteredRows.length > 0) {
                // build arrays for each set
                const data1 = filteredRows.map(row => {
                    const newRow = {};
                    actualSet1Cols.forEach(col => {
                        let val = row[col] || '';
                        // convert potential dates
                        if (typeof val === 'number' || col.toLowerCase().includes('date') || col.toLowerCase().includes('remind')) {
                            val = formatDateValue(val);
                        }
                        newRow[col] = val;
                    });
                    return newRow;
                });
                const data2 = filteredRows.map(row => {
                    const newRow = {};
                    actualSet2Cols.forEach(col => {
                        let val = row[col] || '';
                        if (typeof val === 'number' || col.toLowerCase().includes('date')) {
                            val = formatDateValue(val);
                        }
                        newRow[col] = val;
                    });
                    return newRow;
                });

                window.ktdkSplitResults.set1.dateMap[displayDate] = data1;
                window.ktdkSplitResults.set2.dateMap[displayDate] = data2;
                processedCount++;
            }
        });

        // Hide the date selection container
        document.getElementById('ktdkTypeA_DatesContainer').style.display = 'none';

        // Show results UI
        displaySplitResults(processedCount);
        document.getElementById('ktdkTypeA_ResultsContainer').style.display = 'block';
        document.getElementById('statusKTDKSelect').innerText = `Đã tách ${processedCount} ngày`;
    } catch (err) {
        console.error(err);
        alert('Lỗi khi tách: ' + err.message);
        document.getElementById('statusKTDKSelect').innerText = '';
    }
}

// Display split results with export options
function displaySplitResults(countDates) {
    // build a list of dates for display
    const dates = Object.keys(window.ktdkSplitResults.set1.dateMap);
    let html = `<div style="margin-bottom:15px;"><strong>Danh sách ngày đã tách (${dates.length}):</strong> ${dates.join(', ')}</div>`;

    function sectionHtml(label, setIdx) {
        const is3col = setIdx === 1;
        const fileBase = is3col ? 'KTDK_3COL' : 'KTDK_6COL';
        const desc = is3col ? '3 cột (Thông tin KH)' : '6 cột (Nội dung KTĐK)';
        return `
            <div style="border:1px solid #ddd; padding:10px; margin-bottom:10px; border-radius:5px;">
                <strong>${label}</strong> (${desc})<br>
                <button onclick="showExportDaySelection(${setIdx})" style="margin:5px;">📁 Xuất từng ngày</button>
                <button onclick="exportSplit(${setIdx}, true)" style="margin:5px;">📦 Xuất tổng hợp</button>
            </div>
        `;
    }

    html += sectionHtml('3COL', 1);
    html += sectionHtml('6COL', 2);

    document.getElementById('kdkFilesList').innerHTML = html;
}


// Show dialog to select which days to export
function showExportDaySelection(setIndex) {
    const result = window.ktdkSplitResults && window.ktdkSplitResults[`set${setIndex}`];
    if (!result) {
        alert('Chưa có dữ liệu để xuất');
        return;
    }

    const dates = Object.keys(result.dateMap);
    const is3col = setIndex === 1;
    const fileBase = is3col ? 'KTDK_3COL' : 'KTDK_6COL';

    // Create modal HTML
    let modalHtml = `
        <div style="position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); display:flex; align-items:center; justify-content:center; z-index:9999;" id="exportModal">
            <div style="background:white; padding:20px; border-radius:8px; max-width:400px; max-height:80vh; overflow-y:auto;">
                <h3 style="margin-top:0;">Chọn ngày để xuất ${fileBase}</h3>
                <div style="margin-bottom:15px;">
                    <label style="display:block; margin-bottom:10px; cursor:pointer;">
                        <input type="checkbox" id="exportSelectAllDays">
                        <strong>Chọn tất cả</strong>
                    </label>
    `;

    dates.forEach(dt => {
        modalHtml += `
                    <label style="display:block; margin-bottom:8px; cursor:pointer;">
                        <input type="checkbox" class="exportDateCheck" value="${dt}">
                        ${dt}
                    </label>
        `;
    });

    modalHtml += `
                </div>
                <div style="text-align:right;">
                    <button onclick="closeExportModal()" style="padding:8px 15px; margin-right:10px;">Huỷ</button>
                    <button onclick="doExportSelectedDays(${setIndex})" style="padding:8px 15px; background:#28a745; color:white; border:none; border-radius:4px; cursor:pointer;">Xuất</button>
                </div>
            </div>
        </div>
    `;

    // Inject modal into page
    const modalDiv = document.createElement('div');
    modalDiv.innerHTML = modalHtml;
    document.body.appendChild(modalDiv);

    // Add select-all handler
    document.getElementById('exportSelectAllDays').addEventListener('change', function () {
        const checked = this.checked;
        document.querySelectorAll('.exportDateCheck').forEach(cb => cb.checked = checked);
    });
}

// Close export modal
function closeExportModal() {
    const modal = document.getElementById('exportModal');
    if (modal) modal.parentNode.removeChild(modal);
}

// Export selected days
function doExportSelectedDays(setIndex) {
    const checkboxes = document.querySelectorAll('.exportDateCheck:checked');
    if (checkboxes.length === 0) {
        alert('Vui lòng chọn ít nhất một ngày!');
        return;
    }

    const result = window.ktdkSplitResults && window.ktdkSplitResults[`set${setIndex}`];
    if (!result) {
        alert('Chưa có dữ liệu để xuất');
        return;
    }

    const is3col = setIndex === 1;
    const filePrefix = is3col ? 'KTDK_3COL_ThongTinKH' : 'KTDK_6COL_NoiDungKTDK';
    const dateMap = result.dateMap;

    try {
        checkboxes.forEach(checkbox => {
            const displayDate = checkbox.value; // format: DD-MM-YYYY
            const rows = dateMap[displayDate] ? dateMap[displayDate].length : 0;
            const sheetDay = displayDate.split('-')[0]; // extract day only (e.g., "01")
            const sheetName = `${sheetDay} (${rows})`;
            const wb = XLSX.utils.book_new();

            // Create sheet with headers first, then add data
            const ws = {};

            // Get headers - either from source or create fallback
            let headerRows;
            if (window.ktdkHeaderRows && window.ktdkHeaderRows.length >= 2) {
                headerRows = window.ktdkHeaderRows.slice(0, 2);
            } else {
                // Create fallback headers
                const cols = Object.keys(dateMap[displayDate][0] || {});
                headerRows = [
                    cols, // Row 1: same as row 2 if no original headers
                    cols  // Row 2: column names
                ];
            }

            // Add header rows
            XLSX.utils.sheet_add_aoa(ws, headerRows, { origin: 'A1' });
            // Add data rows starting at row 3 (index 2) WITHOUT generating an extra header row
            if (dateMap[displayDate] && dateMap[displayDate].length) {
                XLSX.utils.sheet_add_json(ws, dateMap[displayDate], { origin: 'A3', skipHeader: true });
            }
            // compute last row index: headerRows.length (2) + data.length - 1 (0-based)
            const _len1 = (dateMap[displayDate] ? dateMap[displayDate].length : 0);
            ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: (_len1 + headerRows.length - 1), c: 20 } });

            XLSX.utils.book_append_sheet(wb, ws, sheetName);
            XLSX.writeFile(wb, `${filePrefix}_${displayDate}.xlsx`);
        });
        alert(`Đã xuất ${checkboxes.length} file thành công!`);
        closeExportModal();
    } catch (err) {
        console.error(err);
        alert('Lỗi khi xuất: ' + err.message);
    }
}

// Export split results for a given set (1=contact,2=test)
function exportSplit(setIndex, consolidated) {
    const result = window.ktdkSplitResults && window.ktdkSplitResults[`set${setIndex}`];
    if (!result) {
        alert('Chưa có dữ liệu để xuất');
        return;
    }

    const is3col = setIndex === 1;
    const filePrefix = is3col ? 'KTDK_3COL_ThongTinKH' : 'KTDK_6COL_NoiDungKTDK';
    const dateMap = result.dateMap;
    const dates = Object.keys(dateMap);

    try {
        if (consolidated) {
            // Extract month-year from first date (format: DD-MM-YYYY)
            const firstDate = dates[0];
            const [day, month, year] = firstDate.split('-');
            const monthYear = `${month}-${year}`;

            const wb = XLSX.utils.book_new();
            dates.forEach(dt => {
                const rows = dateMap[dt] ? dateMap[dt].length : 0;
                const sheetDay = dt.split('-')[0]; // sheet name = day only (e.g., "01")
                const sheetName = `${sheetDay} (${rows})`;

                // Create sheet with headers first, then add data
                const ws = {};

                // Get headers - either from source or create fallback
                let headerRows;
                if (window.ktdkHeaderRows && window.ktdkHeaderRows.length >= 2) {
                    headerRows = window.ktdkHeaderRows.slice(0, 2);
                } else {
                    // Create fallback headers
                    const cols = Object.keys(dateMap[dt][0] || {});
                    headerRows = [
                        cols, // Row 1: same as row 2 if no original headers
                        cols  // Row 2: column names
                    ];
                }

                // Add header rows
                XLSX.utils.sheet_add_aoa(ws, headerRows, { origin: 'A1' });
                // Add data rows starting at row 3 (index 2) WITHOUT generating an extra header row
                if (dateMap[dt] && dateMap[dt].length) {
                    XLSX.utils.sheet_add_json(ws, dateMap[dt], { origin: 'A3', skipHeader: true });
                }
                const _len2 = (dateMap[dt] ? dateMap[dt].length : 0);
                ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: (_len2 + headerRows.length - 1), c: 20 } });

                XLSX.utils.book_append_sheet(wb, ws, sheetName);
            });
            XLSX.writeFile(wb, `${filePrefix}_${monthYear}.xlsx`);
            alert(`Đã xuất file tổng hợp ${filePrefix}_${monthYear}.xlsx`);
        }
    } catch (err) {
        console.error(err);
        alert('Lỗi khi xuất: ' + err.message);
    }
}

// ========== TYPE B: Import Standard Format ==========
// Type B: Read template file structure
async function readTemplateFile() {
    const fileInput = document.getElementById('kdkTemplateFile');

    if (!fileInput.files || fileInput.files.length === 0) {
        alert('Vui lòng chọn file mẫu!');
        return;
    }

    try {
        document.getElementById('statusKTDKTemplate').innerText = 'Đang đọc file mẫu...';

        // read raw rows so we can inspect header lines
        const dataRaw = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                const firstSheet = wb.Sheets[wb.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
                resolve(rows);
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(fileInput.files[0]);
        });

        if (!dataRaw || dataRaw.length === 0) {
            alert('File mẫu không có dữ liệu!');
            document.getElementById('statusKTDKTemplate').innerText = '';
            return;
        }

        // choose header row: by default first, but if second row seems more complete, use it
        let headerRow = dataRaw[0];
        if (dataRaw.length > 1) {
            const second = dataRaw[1];
            // pick second if it has more non-empty cells
            const count1 = headerRow.filter(c => c !== '').length;
            const count2 = second.filter(c => c !== '').length;
            if (count2 > count1) {
                headerRow = second;
            }
            // also if second contains any key words like MODEL_NAME, TEST_TYPE
            const keywords = ['MODEL_NAME', 'TEST_TYPE', 'CONTENT', 'ENDDATE', 'PLACENAME', 'NUMBER_PLATE'];
            if (keywords.some(k => second.join(' ').toUpperCase().includes(k))) {
                headerRow = second;
            }
        }

        const templateColumns = headerRow.map(c => String(c).trim()).filter(c => c !== '');

        // store globally for later use
        window.kdkTemplateColumns = templateColumns;
        window.kdkTemplateFile = fileInput.files[0];

        // build display html for columns
        const colsHtml = templateColumns.map((col, idx) =>
            `<span style="display:inline-block; margin:3px 8px; padding:4px 8px; background:#e3f2fd; border-radius:3px; font-size:11px;">
                <strong>${idx + 1}.</strong> ${col}
            </span>`
        ).join('');

        document.getElementById('kdkTemplateColumns').innerHTML = colsHtml;
        document.getElementById('kdkTemplateInfo').style.display = 'block';
        document.getElementById('kdkTypeB_SettingsAfterTemplate').style.display = 'block';

        let statusMsg = `✓ Tải thành công file mẫu (${templateColumns.length} cột)`;
        if (headerRow !== dataRaw[0]) {
            statusMsg += ' (đã sử dụng dòng 2 làm tiêu đề)';
        }
        document.getElementById('statusKTDKTemplate').innerText = statusMsg;

    } catch (err) {
        console.error(err);
        alert('Lỗi khi đọc file mẫu: ' + err.message);
        document.getElementById('statusKTDKTemplate').innerText = '';
    }
}

async function selectDatesToSplitKTDKTypeB() {
    const dateColName = document.getElementById('kdkDateColumnB').value.trim();
    const monthYear = document.getElementById('kdkMonthYearB').value.trim();

    if (!window.kdkTemplateColumns || !window.kdkTemplateColumns.length) {
        alert('Vui lòng tải file mẫu trước!');
        return;
    }

    if (!window.ktdkData) {
        alert('Vui lòng đọc file dữ liệu trước!');
        return;
    }

    if (!dateColName) {
        alert('Vui lòng nhập tên cột ngày!');
        return;
    }

    if (!monthYear) {
        alert('Vui lòng nhập tháng-năm (ví dụ: T03_2026)!');
        return;
    }

    const data = window.ktdkData;

    // Check column exists
    if (!(dateColName in data[0])) {
        alert(`Không tìm thấy cột "${dateColName}" trong file!`);
        return;
    }

    try {
        document.getElementById('statusKTDKSelectB').innerText = 'Đang tải danh sách ngày...';

        // Get unique dates with proper formatting
        const uniqueDates = new Set();
        data.forEach(row => {
            const dateVal = row[dateColName];
            if (dateVal) {
                const formatted = formatDateValue(dateVal);
                uniqueDates.add(formatted);
            }
        });

        const sortedDates = Array.from(uniqueDates).sort();

        // Create pairs of dates (day 1-2, 3-4, 5-6, etc.)
        let html = '';
        for (let i = 0; i < sortedDates.length; i += 2) {
            const date1 = sortedDates[i];
            const date2 = i + 1 < sortedDates.length ? sortedDates[i + 1] : null;
            const rows1 = data.filter(r => formatDateValue(r[dateColName]) === date1).length;
            const rows2 = date2 ? data.filter(r => formatDateValue(r[dateColName]) === date2).length : 0;
            const totalRows = rows1 + (date2 ? rows2 : 0);
            const pairLabel = date2 ? `${date1} - ${date2}` : `${date1}`;
            const pairValue = date2 ? `${date1},${date2}` : date1;

            html += `<label style="display:inline-block; margin:10px 20px 10px 0; cursor:pointer;">
                <input type="checkbox" class="kdkDateCheckB" value="${pairValue}">
                ${pairLabel} (${totalRows} dòng)
            </label>`;
        }

        document.getElementById('kdkDatesListB').innerHTML = html;
        // add select-all handler
        document.getElementById('kdkDateSelectAllB').addEventListener('change', function () {
            const checked = this.checked;
            document.querySelectorAll('.kdkDateCheckB').forEach(cb => cb.checked = checked);
        });

        document.getElementById('ktdkTypeB_DatesContainer').style.display = 'block';
        document.getElementById('statusKTDKSelectB').innerText = `Tìm thấy ${Math.ceil(sortedDates.length / 2)} cặp ngày`;

        // Store column name and month-year globally
        window.kdkDateColumnNameB = dateColName;
        window.kdkMonthYearB = monthYear;

    } catch (err) {
        console.error(err);
        alert('Lỗi: ' + err.message);
        document.getElementById('statusKTDKSelectB').innerText = '';
    }
}

async function doSplitKTDKTypeB() {
    const checkboxes = document.querySelectorAll('.kdkDateCheckB:checked');
    if (checkboxes.length === 0) {
        alert('Vui lòng chọn ít nhất một cặp ngày!');
        return;
    }

    if (!window.ktdkData || !window.kdkDateColumnNameB || !window.kdkMonthYearB || !window.kdkTemplateColumns) {
        alert('Vui lòng tải file mẫu, đọc dữ liệu và chọn cặp ngày trước!');
        return;
    }

    try {
        document.getElementById('statusKTDKSelectB').innerText = 'Đang tạo file theo mẫu chuẩn...';

        const data = window.ktdkData;
        const dateColName = window.kdkDateColumnNameB;
        const monthYear = window.kdkMonthYearB;
        const templateColumns = window.kdkTemplateColumns;

        // Helper to find column name case-insensitively
        function findColumnName(searchNames) {
            if (!data || data.length === 0) return null;
            const headers = Object.keys(data[0]);
            for (let searchName of searchNames) {
                const found = headers.find(h => h.toLowerCase() === searchName.toLowerCase());
                if (found) return found;
            }
            return null;
        }

        // Create mapping from template column names to source column names
        const columnMapping = {};
        templateColumns.forEach(templateCol => {
            const templateColLower = templateCol.toLowerCase();

            // Try to find matching source column (case-insensitive)
            let sourceCol = null;

            // Direct match
            if (templateColLower === 'fullname') {
                sourceCol = findColumnName(['FULLNAME', 'FullName', 'Full_Name']);
            } else if (templateColLower === 'customer_name' || templateColLower === 'customername') {
                // customer_name template should be filled from either the
                // explicit CUSTOMER_NAME column or the FullName column as a fallback
                sourceCol = findColumnName(['CUSTOMER_NAME', 'CustomerName', 'Customer_Name', 'FULLNAME', 'FullName', 'Full_Name']);
            } else if (templateColLower === 'servicetype' || templateColLower === 'service_type' || templateColLower === 'service type') {
                // always default ServiceType to 'KTDK' if not present in source data
                sourceCol = findColumnName(['ServiceType', 'SERVICETYPE', 'service_type']) || null;
            } else if (templateColLower === 'phoneno') {
                sourceCol = findColumnName(['TELEINFO', 'TeleInfo', 'PHONE', 'MOBILE', 'PhoneNo']);
            } else if (templateColLower === 'branchcode') {
                sourceCol = findColumnName(['BRANCHCODE', 'BranchCode', 'Branch_Code', 'BRANCH_CODE', 'Branch']);
            } else if (templateColLower === 'placename') {
                sourceCol = findColumnName(['PLACE', 'Place', 'PLACENAME']);
            } else {
                // Try exact match or case-insensitive match
                const headers = Object.keys(data[0]);
                sourceCol = headers.find(h => h.toLowerCase() === templateColLower);
            }

            columnMapping[templateCol] = sourceCol;
        });

        // Process each selected date pair
        const filesToCreate = [];

        checkboxes.forEach((checkbox) => {
            const pairValue = checkbox.value; // "YYYY-MM-DD,YYYY-MM-DD" or single date
            const pairDates = pairValue.split(',');

            // Filter data for these dates
            const filteredRows = data.filter(row => {
                const rowDate = formatDateValue(row[dateColName]);
                return pairDates.includes(rowDate);
            });

            if (filteredRows.length === 0) return;

            // Map columns to template format
            const mappedRows = filteredRows.map(sourceRow => {
                const templateRow = {};
                templateColumns.forEach(templateCol => {
                    const sourceCol = columnMapping[templateCol];
                    let value = sourceCol ? (sourceRow[sourceCol] || '') : '';

                    const templLower = templateCol.toLowerCase();

                    // if this is a ServiceType column and we don't have a value,
                    // default to 'KTDK'
                    if ((templLower === 'servicetype' || templLower === 'service_type' || templLower === 'service type') && (!value || String(value).trim() === '')) {
                        value = 'KTDK';
                    }

                    // Format dates if applicable
                    if ((templLower.includes('date')) && value) {
                        value = formatDateValue(value);
                    }

                    templateRow[templateCol] = value;
                });
                return templateRow;
            });

            // Create filename
            // Format: IMPORT KTDK T03_2026 _ RemindDate 01-02 _ 957
            let dateRange = pairDates[0].split('-')[2]; // day from first date
            if (pairDates.length > 1) {
                dateRange += '-' + pairDates[1].split('-')[2]; // day from second date
            }
            const rowCount = mappedRows.length;
            const filename = `IMPORT KTDK ${monthYear} _ RemindDate ${dateRange} _ ${rowCount}.xlsx`;

            filesToCreate.push({
                filename: filename,
                data: mappedRows,
                columns: templateColumns
            });
        });

        // Export files
        if (filesToCreate.length === 0) {
            alert('Không có dữ liệu để xuất!');
            document.getElementById('statusKTDKSelectB').innerText = '';
            return;
        }

        filesToCreate.forEach(({ filename, data, columns }) => {
            const wb = XLSX.utils.book_new();

            // Create sheet with headers first, then add data
            const ws = {};

            // Get headers - either from source or create fallback
            let headerRows;
            if (window.ktdkHeaderRows && window.ktdkHeaderRows.length >= 2) {
                headerRows = window.ktdkHeaderRows.slice(0, 2);
            } else {
                // Create fallback headers from columns
                headerRows = [
                    columns, // Row 1: same as row 2 if no original headers
                    columns  // Row 2: column names
                ];
            }

            // Add header rows
            XLSX.utils.sheet_add_aoa(ws, headerRows, { origin: 'A1' });
            // Add data rows starting at row 3 (index 2) WITHOUT generating an extra header row
            if (data && data.length) {
                XLSX.utils.sheet_add_json(ws, data, { origin: 'A3', skipHeader: true });
            }
            const _len3 = (data ? data.length : 0);
            ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: (_len3 + headerRows.length - 1), c: 20 } });

            // sheet name set to number of rows
            const sheetName = `${data.length}`;
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
            XLSX.writeFile(wb, filename);
        });

        // Hide date selection and show results
        document.getElementById('ktdkTypeB_DatesContainer').style.display = 'none';

        // Show results
        let resultHtml = `<div style="margin-bottom:15px;"><strong>Đã xuất ${filesToCreate.length} file:</strong></div>`;
        filesToCreate.forEach(({ filename }) => {
            resultHtml += `<div style="padding:8px; background:#f0f0f0; margin:5px 0; border-radius:3px;">✓ ${filename}</div>`;
        });

        document.getElementById('kdkFilesListB').innerHTML = resultHtml;
        document.getElementById('ktdkTypeB_ResultsContainer').style.display = 'block';
        document.getElementById('statusKTDKSelectB').innerText = `Đã xuất ${filesToCreate.length} file thành công!`;
    } catch (err) {
        console.error(err);
        alert('Lỗi khi xuất file: ' + err.message);
        document.getElementById('statusKTDKSelectB').innerText = '';
    }
}

// initialize with compare section visible
window.addEventListener('DOMContentLoaded', () => {
    showSection('compare');
});
