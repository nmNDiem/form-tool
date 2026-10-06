import * as XLSX from 'xlsx';

const wait = (ms) => new Promise(resolve => setTimeout(resolve, ms));

export const analyzeExcel = async (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        if (jsonData.length === 0) {
          throw new Error("File Excel trống hoặc không đọc được dữ liệu.");
        }
        
        const headers = jsonData[0];
        const existingDataRows = [];
        for (let i = 1; i < jsonData.length; i++) {
          const row = jsonData[i];
          if (row && row.length > 0 && row.some(cell => cell !== undefined && cell !== null && cell !== '')) {
            existingDataRows.push(row);
          }
        }
        
        const currentRowsCount = existingDataRows.length;
        
        const standardColIndices = [];
        let q2Index = -1;

        for (let i = 0; i < headers.length; i++) {
          const headerName = headers[i]?.toString() || '';
          // Check for Q2 column
          if (headerName.includes('Q2') || headerName === 'Q2') {
             q2Index = i;
          }

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
        
        const analysisResult = {
          headers,
          existingDataRows,
          standardColIndices,
          currentRowsCount,
          firstSheetName,
          hasQ2: q2Index !== -1,
          q2Index,
          q2Groups: {},
          overallDistributions: {},
          stats: {
            originalRows: currentRowsCount,
            columnsProcessed: standardColIndices.length,
            columnStats: {}
          }
        };

        const calculateDistribution = (rows, colIndices) => {
          const dists = {};
          const colStats = {};
          for (const colIndex of colIndices) {
            const colName = headers[colIndex] || `Cột ${colIndex + 1}`;
            const counts = {};
            let initialTotalValid = 0;
            
            for (const row of rows) {
              let val = row[colIndex];
              if (val !== undefined && val !== null && val !== '') {
                 val = val.toString().trim();
                 counts[val] = (counts[val] || 0) + 1;
                 initialTotalValid++;
              }
            }
            
            let filteredTotalValid = 0;
            const filteredCounts = {};
            
            for (const [val, count] of Object.entries(counts)) {
              const percentage = count / initialTotalValid;
              if (percentage >= 0.1) {
                 filteredCounts[val] = count;
                 filteredTotalValid += count;
              }
            }
            
            colStats[colName] = { 
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
                  distArray.push({ value: val, cumulativeProb, prob });
              }
              dists[colIndex] = distArray;
            } else {
               dists[colIndex] = [];
            }
          }
          return { dists, colStats };
        };

        if (q2Index !== -1) {
          const rowsByQ2 = {};
          for (const row of existingDataRows) {
             let q2Val = row[q2Index];
             if (q2Val !== undefined && q2Val !== null && q2Val !== '') {
               q2Val = q2Val.toString().trim();
               if (!rowsByQ2[q2Val]) rowsByQ2[q2Val] = [];
               rowsByQ2[q2Val].push(row);
             }
          }
          
          const q2Result = calculateDistribution(existingDataRows, [q2Index]);
          analysisResult.overallDistributions[q2Index] = q2Result.dists[q2Index];
          analysisResult.stats.columnStats['Tổng quan'] = q2Result.colStats;
          
          const otherColIndices = standardColIndices.filter(i => i !== q2Index);
          for (const [q2Val, rows] of Object.entries(rowsByQ2)) {
             const groupResult = calculateDistribution(rows, otherColIndices);
             analysisResult.q2Groups[q2Val] = {
               count: rows.length,
               distributions: groupResult.dists,
               stats: groupResult.colStats
             };
          }
        } else {
          const result = calculateDistribution(existingDataRows, standardColIndices);
          analysisResult.overallDistributions = result.dists;
          analysisResult.stats.columnStats['Tổng quan'] = result.colStats;
        }
        
        resolve(analysisResult);
        
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (error) => reject(new Error("Lỗi đọc file: " + error.message));
    reader.readAsArrayBuffer(file);
  });
};

export const generateExcel = async (analysisResult, targetRows, onProgress) => {
  const log = async (msg) => {
    if (onProgress) onProgress(msg);
    await wait(100);
  };
  
  try {
    const { headers, existingDataRows, standardColIndices, currentRowsCount, firstSheetName, hasQ2, q2Index, q2Groups, overallDistributions } = analysisResult;
    
    if (currentRowsCount >= targetRows) {
      throw new Error(`File đã có ${currentRowsCount} dòng dữ liệu, lớn hơn hoặc bằng số lượng mong muốn (${targetRows}). Không cần thêm.`);
    }
    
    const rowsToAdd = targetRows - currentRowsCount;
    await log(`Chuẩn bị tự động điền ngẫu nhiên ${rowsToAdd} dòng mới...`);

    const newRows = [];
    for (let r = 0; r < rowsToAdd; r++) {
      const newRow = new Array(headers.length).fill('');
      
      if (hasQ2) {
         const q2Dist = overallDistributions[q2Index];
         let selectedQ2 = '';
         if (q2Dist && q2Dist.length > 0) {
             const rand = Math.random();
             selectedQ2 = q2Dist[q2Dist.length - 1].value; 
             for (const item of q2Dist) {
               if (rand <= item.cumulativeProb) {
                 selectedQ2 = item.value;
                 break;
               }
             }
         }
         newRow[q2Index] = selectedQ2;
         
         const group = q2Groups[selectedQ2];
         for (const colIndex of standardColIndices) {
            if (colIndex === q2Index) continue;
            
            const dist = group && group.distributions ? group.distributions[colIndex] : [];
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
      } else {
         for (const colIndex of standardColIndices) {
            const dist = overallDistributions[colIndex];
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
      }
      
      newRows.push(newRow);
      
      if (r > 0 && r % 100 === 0) {
        await log(` Đã điền xong ${r}/${rowsToAdd} dòng...`);
      }
    }
    
    await log(`Hoàn tất việc tạo ${rowsToAdd} dòng ngẫu nhiên. Đang đóng gói file...`);

    const finalData = [headers, ...existingDataRows, ...newRows];
    
    const newWorksheet = XLSX.utils.aoa_to_sheet(finalData);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, firstSheetName || "Sheet1");
    
    const excelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
    
    await log(`Ghi file Excel thành công!`);

    return {
      buffer: excelBuffer,
      stats: {
         originalRows: currentRowsCount,
         addedRows: rowsToAdd,
         totalRows: targetRows,
      }
    };
  } catch (err) {
      throw err;
  }
};
