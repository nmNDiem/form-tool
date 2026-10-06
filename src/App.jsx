import React, { useState, useRef, useEffect } from 'react';
import { UploadCloud, FileSpreadsheet, X, AlertCircle, CheckCircle, Download, FilePlus2, Loader2, ChevronRight, BarChart2 } from 'lucide-react';
import { analyzeExcel, generateExcel } from './utils/excelProcessor';

function App() {
  const [file, setFile] = useState(null);
  const [targetRows, setTargetRows] = useState(300);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [analysisResult, setAnalysisResult] = useState(null);
  const [logs, setLogs] = useState([]);
  const [error, setError] = useState(null);
  const [result, setResult] = useState(null);
  const fileInputRef = useRef(null);
  const logsEndRef = useRef(null);

  useEffect(() => {
    if (logsEndRef.current) {
      logsEndRef.current.scrollIntoView({ behavior: 'smooth' });
    }
  }, [logs]);

  const handleDragOver = (e) => {
    e.preventDefault();
    e.currentTarget.classList.add('drag-active');
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-active');
  };

  const handleFileChange = async (selectedFile) => {
    if (selectedFile) {
      setFile(selectedFile);
      setError(null);
      setResult(null);
      setLogs([]);
      setAnalysisResult(null);
      
      setIsAnalyzing(true);
      try {
        const res = await analyzeExcel(selectedFile);
        setAnalysisResult(res);
        setTargetRows(Math.max(300, res.currentRowsCount + 100));
      } catch (err) {
        setError(err.message || 'Có lỗi xảy ra khi phân tích file');
      } finally {
        setIsAnalyzing(false);
      }
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-active');
    const droppedFile = e.dataTransfer.files[0];
    if (droppedFile && (droppedFile.name.endsWith('.xlsx') || droppedFile.name.endsWith('.xls'))) {
      handleFileChange(droppedFile);
    } else {
      setError('Vui lòng chọn file Excel (.xlsx hoặc .xls)');
    }
  };

  const clearFile = () => {
    setFile(null);
    setResult(null);
    setError(null);
    setLogs([]);
    setAnalysisResult(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleGenerate = async () => {
    if (!analysisResult) return;
    if (targetRows <= analysisResult.currentRowsCount) {
      setError(`Số dòng phải lớn hơn số dòng hiện tại (${analysisResult.currentRowsCount})`);
      return;
    }

    setIsGenerating(true);
    setLogs([]);
    setError(null);
    
    try {
      const res = await generateExcel(analysisResult, parseInt(targetRows, 10), (msg) => {
        setLogs(prev => [...prev, msg]);
      });
      setResult(res);
    } catch (err) {
      setError(err.message || 'Có lỗi xảy ra trong quá trình xử lý');
      setLogs(prev => [...prev, `[LỖI] ${err.message}`]);
    } finally {
      setIsGenerating(false);
    }
  };

  const downloadFile = () => {
    if (!result || !result.buffer) return;
    
    const blob = new Blob([result.buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Kết_quả_${file.name}`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="glass-card fade-in">
      <div className="header">
        <h1>AutoFill Excel Options</h1>
        <p>Sinh ngẫu nhiên dữ liệu theo tỉ lệ (đã lọc &lt;10%)</p>
      </div>

      {!file ? (
        <div 
          className="upload-zone"
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
        >
          <UploadCloud className="upload-icon" />
          <div>
            <h3>Kéo thả file Excel vào đây</h3>
            <p style={{ color: 'var(--text-muted)', marginTop: '0.5rem', fontSize: '0.9rem' }}>
              hoặc click để chọn file (.xlsx, .xls)
            </p>
          </div>
          <input 
            type="file" 
            ref={fileInputRef} 
            onChange={(e) => handleFileChange(e.target.files[0])} 
            accept=".xlsx, .xls" 
            className="file-input" 
          />
        </div>
      ) : (
        <div className="fade-in">
          <div className="file-info">
            <div className="file-name">
              <FileSpreadsheet size={24} color="var(--primary)" />
              {file.name}
            </div>
            {!isGenerating && !isAnalyzing && (
              <button className="remove-file" onClick={clearFile} title="Xóa file">
                <X size={20} />
              </button>
            )}
          </div>

          {isAnalyzing && (
             <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', margin: '2rem 0', color: 'var(--primary)' }}>
               <Loader2 size={40} className="loader" style={{ animation: 'spin 1.5s linear infinite', marginBottom: '1rem' }} />
               <h4>Đang phân tích file...</h4>
             </div>
          )}

          {error && !isGenerating && (
            <div className="alert alert-error" style={{ marginBottom: '1rem' }}>
              <AlertCircle size={20} />
              <span>{error}</span>
            </div>
          )}

          {analysisResult && !result && !isGenerating && (
            <div className="fade-in">
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '1rem', color: 'var(--text-main)' }}>
                <BarChart2 size={24} color="var(--primary)" />
                <h3 style={{ margin: 0 }}>Kết quả phân tích</h3>
              </div>
              <p style={{ marginBottom: '1rem', fontSize: '0.95rem' }}>Số dòng hợp lệ hiện tại: <strong>{analysisResult.currentRowsCount}</strong></p>
              
              <div className="analysis-details" style={{ maxHeight: '350px', overflowY: 'auto', background: 'rgba(255,255,255,0.4)', padding: '1rem', borderRadius: '12px', marginBottom: '1.5rem', border: '1px solid var(--card-border)' }}>
                {analysisResult.hasQ2 ? (
                   <div>
                      <h4 style={{ color: 'var(--primary)', marginBottom: '1rem', fontSize: '1rem' }}>
                        Đã tìm thấy cột '{analysisResult.headers[analysisResult.q2Index]}' - Phân chia tỉ lệ theo từng nhóm:
                      </h4>
                      {/* Overall Q2 distribution */}
                      <div style={{ marginBottom: '1rem', paddingBottom: '1rem', borderBottom: '2px solid var(--card-border)' }}>
                        <h5 style={{ fontSize: '0.95rem', marginBottom: '0.5rem', color: 'var(--text-main)' }}>Tỉ lệ xuất hiện của '{analysisResult.headers[analysisResult.q2Index]}':</h5>
                        <div style={{ fontSize: '0.9rem', marginLeft: '1rem' }}>
                          {analysisResult.overallDistributions[analysisResult.q2Index]?.map(d => `${d.value} (${(d.prob * 100).toFixed(1)}%)`).join(' | ')}
                        </div>
                      </div>

                      {Object.entries(analysisResult.q2Groups).map(([q2Val, group]) => (
                         <div key={q2Val} style={{ marginBottom: '1.5rem', paddingBottom: '1rem', borderBottom: '1px dashed var(--card-border)' }}>
                           <h5 style={{ fontSize: '0.95rem', marginBottom: '0.8rem', color: 'var(--text-main)' }}>▶ Nhóm {analysisResult.headers[analysisResult.q2Index]} = <strong>{q2Val}</strong> ({group.count} dòng)</h5>
                           {Object.entries(group.distributions).map(([colIndex, dist]) => {
                              const colName = analysisResult.headers[colIndex];
                              if (!dist || dist.length === 0) return null;
                              return (
                                 <div key={colIndex} style={{ fontSize: '0.9rem', marginLeft: '1.5rem', marginBottom: '0.5rem' }}>
                                    <span style={{ color: 'var(--text-muted)' }}>{colName}:</span> <strong>{dist.map(d => `${d.value} (${(d.prob * 100).toFixed(1)}%)`).join(' | ')}</strong>
                                 </div>
                              );
                           })}
                         </div>
                      ))}
                   </div>
                ) : (
                   <div>
                      <h4 style={{ color: 'var(--primary)', marginBottom: '1rem', fontSize: '1rem' }}>Tỉ lệ phân bổ tổng quan (Không có cột Q2):</h4>
                      {Object.entries(analysisResult.overallDistributions).map(([colIndex, dist]) => {
                          const colName = analysisResult.headers[colIndex];
                          if (!dist || dist.length === 0) return null;
                          return (
                             <div key={colIndex} style={{ fontSize: '0.9rem', marginBottom: '0.8rem', marginLeft: '0.5rem' }}>
                                <span style={{ color: 'var(--text-muted)' }}>{colName}:</span> <strong>{dist.map(d => `${d.value} (${(d.prob * 100).toFixed(1)}%)`).join(' | ')}</strong>
                             </div>
                          );
                       })}
                   </div>
                )}
              </div>

              <div className="form-group" style={{ background: 'var(--card-bg)', padding: '1rem', borderRadius: '12px', border: '1px solid var(--primary)', opacity: 0.8 }}>
                <label style={{ color: 'var(--primary)', fontWeight: 600 }}>Nhập tổng số dòng dữ liệu mong muốn (Bao gồm cả cũ và mới):</label>
                <input 
                  type="number" 
                  className="input-field"
                  value={targetRows}
                  onChange={(e) => setTargetRows(e.target.value)}
                  min={analysisResult.currentRowsCount + 1}
                  style={{ border: '2px solid var(--primary)' }}
                />
              </div>

              <button 
                className="btn" 
                style={{ width: '100%', justifyContent: 'center', marginTop: '1.5rem', padding: '1rem', background: 'var(--success)' }}
                onClick={handleGenerate}
              >
                <CheckCircle size={22} />
                Đồng ý điền dữ liệu
              </button>
            </div>
          )}

          {(isGenerating || logs.length > 0) && (
            <div className="logs-container" style={{ margin: '1.5rem 0', background: 'rgba(255,255,255,0.5)', borderRadius: '12px', padding: '1rem', border: '1px solid var(--card-border)', maxHeight: '250px', overflowY: 'auto' }}>
              <h4 style={{ marginBottom: '0.8rem', color: 'var(--text-main)', fontSize: '0.95rem' }}>Tiến trình xử lý:</h4>
              {logs.map((log, index) => (
                <div key={index} style={{ fontSize: '0.9rem', color: log.includes('[LỖI]') ? 'var(--error)' : 'var(--text-muted)', marginBottom: '0.5rem', display: 'flex', alignItems: 'flex-start', gap: '0.5rem' }}>
                  <ChevronRight size={16} style={{ flexShrink: 0, marginTop: '2px', color: log.includes('[LỖI]') ? 'var(--error)' : 'var(--primary)' }} />
                  <span>{log}</span>
                </div>
              ))}
              {isGenerating && (
                 <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', color: 'var(--primary)', fontSize: '0.9rem', marginTop: '1rem' }}>
                    <Loader2 size={16} className="loader" style={{ border: 'none', animation: 'spin 1.5s linear infinite' }} />
                    <span>Đang thực hiện...</span>
                 </div>
              )}
              <div ref={logsEndRef} />
            </div>
          )}

          {result && (
            <div className="fade-in">
              <div style={{ textAlign: 'center', marginBottom: '1.5rem', color: 'var(--success)' }}>
                <CheckCircle size={48} style={{ margin: '0 auto 1rem' }} />
                <h3 style={{ fontSize: '1.2rem' }}>Hoàn tất xử lý!</h3>
              </div>
              
              <div className="stats-grid">
                <div className="stat-card">
                  <div className="stat-value">{result.stats.originalRows}</div>
                  <div className="stat-label">Dòng có sẵn</div>
                </div>
                <div className="stat-card">
                  <div className="stat-value" style={{ color: '#8b5cf6' }}>+{result.stats.addedRows}</div>
                  <div className="stat-label">Dòng sinh thêm</div>
                </div>
                <div className="stat-card">
                  <div className="stat-value" style={{ color: '#10b981' }}>{result.stats.totalRows}</div>
                  <div className="stat-label">Tổng số dòng</div>
                </div>
              </div>

              <div style={{ display: 'flex', gap: '1rem', marginTop: '1.5rem' }}>
                <button className="btn btn-outline" style={{ flex: 1, justifyContent: 'center' }} onClick={clearFile}>
                  Làm lại
                </button>
                <button className="btn" style={{ flex: 2, justifyContent: 'center' }} onClick={downloadFile}>
                  <Download size={20} />
                  Tải file kết quả
                </button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

export default App;
