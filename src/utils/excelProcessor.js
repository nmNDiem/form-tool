import * as XLSX from 'xlsx';

const wait = (ms) => new Promise(resolve => setTimeout(resolve, ms));

export const processExcel = async (file, targetRows, onProgress) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const log = async (msg) => {
          if (onProgress) onProgress(msg);
          await wait(300);
        };

        await log(`Bắt đầu đọc file: ${file.name}`);

        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        if (jsonData.length === 0) {
          throw new Error("File Excel trống hoặc không đọc được dữ liệu.");
        }
        
        const headers = jsonData[0];
        await log(`Đã tải thành công ${headers.length} cột dữ liệu.`);

        const existingDataRows = [];
        for (let i = 1; i < jsonData.length; i++) {
          const row = jsonData[i];
          if (row && row.length > 0 && row.some(cell => cell !== undefined && cell !== null && cell !== '')) {
            existingDataRows.push(row);
          }
        }
        
        const currentRowsCount = existingDataRows.length;
        await log(`Phân tích thấy ${currentRowsCount} dòng dữ liệu hợp lệ (bỏ qua dòng trống).`);

        if (currentRowsCount >= targetRows) {
          throw new Error(`File đã có ${currentRowsCount} dòng dữ liệu, lớn hơn hoặc bằng số lượng mong muốn (${targetRows}). Không cần thêm.`);
        }
        
        await log("Đang quét để xác định các cột cần xử lý (các cột có chứa dữ liệu)...");
        
        const standardColIndices = [];
        for (let i = 0; i < headers.length; i++) {
          // A column is processed if it has at least one non-empty cell in the existing data rows
          let hasData = false;
          for (const row of existingDataRows) {
             const val = row[i];
             if (val !== undefined && val !== null && val !== '') {
               hasData = true;
               break;
             }
          }
          if (hasData) {
            standardColIndices.push(i);
          }
        }
        
        if (standardColIndices.length === 0) {
          throw new Error("Tất cả các cột đều trống. Không có dữ liệu để tính toán tỉ lệ.");
        }
        
        await log(`Phát hiện ${standardColIndices.length} cột có dữ liệu. Bắt đầu tính toán tỉ lệ phân bổ...`);

        const rowsToAdd = targetRows - currentRowsCount;
        const distributions = {}; 
        const stats = {
            originalRows: currentRowsCount,
            addedRows: rowsToAdd,
            totalRows: targetRows,
            columnsProcessed: standardColIndices.length,
            columnStats: {}
        };
        
        for (const colIndex of standardColIndices) {
          const colName = headers[colIndex] || `Cột ${colIndex + 1}`;
          const counts = {};
          let initialTotalValid = 0;
          
          for (const row of existingDataRows) {
            let val = row[colIndex];
            if (val !== undefined && val !== null && val !== '') {
               val = val.toString().trim();
               counts[val] = (counts[val] || 0) + 1;
               initialTotalValid++;
            }
          }
          
          let filteredTotalValid = 0;
          const filteredCounts = {};
          let droppedOptions = [];
          
          for (const [val, count] of Object.entries(counts)) {
            const percentage = count / initialTotalValid;
            if (percentage >= 0.1) {
               filteredCounts[val] = count;
               filteredTotalValid += count;
            } else {
               droppedOptions.push(val);
            }
          }
          
          if (droppedOptions.length > 0) {
            await log(`- Cột "${colName}": Loại bỏ các option dưới 10% (${droppedOptions.join(', ')}).`);
          }

          stats.columnStats[colName] = { 
            originalCounts: counts, 
            filteredCounts,
            initialTotalValid,
            filteredTotalValid
          };
          
          if (filteredTotalValid > 0) {
            const distArray = [];
            let cumulativeProb = 0;
            for (const [val, count] of Object.entries(filteredCounts)) {
                const prob = count / filteredTotalValid;
                cumulativeProb += prob;
                distArray.push({ value: val, cumulativeProb });
            }
            distributions[colIndex] = distArray;
          } else {
             await log(`[Cảnh báo] Cột "${colName}" không còn option nào sau khi lọc <10%. Sẽ không tự điền.`);
             distributions[colIndex] = [];
          }
        }
        
        await log(`Chuẩn bị tự động điền ngẫu nhiên ${rowsToAdd} dòng mới...`);

        const newRows = [];
        for (let r = 0; r < rowsToAdd; r++) {
          const newRow = new Array(headers.length).fill('');
          
          for (const colIndex of standardColIndices) {
             const dist = distributions[colIndex];
             if (!dist || dist.length === 0) continue;

             const rand = Math.random();
             let selectedVal = dist[dist.length - 1].value; 
             for (const item of dist) {
               if (rand <= item.cumulativeProb) {
                 selectedVal = item.value;
                 break;
               }
             }
             newRow[colIndex] = selectedVal;
          }
          newRows.push(newRow);
          
          // Log progress periodically for large datasets
          if (r > 0 && r % 100 === 0) {
            await log(` Đã điền xong ${r}/${rowsToAdd} dòng...`);
          }
        }
        
        await log(`Hoàn tất việc tạo ${rowsToAdd} dòng ngẫu nhiên. Đang đóng gói file...`);

        const finalData = [headers, ...existingDataRows, ...newRows];
        
        const newWorksheet = XLSX.utils.aoa_to_sheet(finalData);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, firstSheetName);
        
        const excelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
        
        await log(`Ghi file Excel thành công!`);

        resolve({
          buffer: excelBuffer,
          stats
        });
        
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (error) => reject(new Error("Lỗi đọc file: " + error.message));
    reader.readAsArrayBuffer(file);
  });
};
