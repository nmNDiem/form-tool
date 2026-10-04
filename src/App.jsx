import React, { useState, useRef, useEffect } from 'react';
import { UploadCloud, FileSpreadsheet, X, AlertCircle, CheckCircle, Download, FilePlus2, Loader2, ChevronRight } from 'lucide-react';
import { processExcel } from './utils/excelProcessor';

function App() {
  const [file, setFile] = useState(null);
  const [targetRows, setTargetRows] = useState(300);
  const [isProcessing, setIsProcessing] = useState(false);
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

  const handleDrop = (e) => {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-active');
    const droppedFile = e.dataTransfer.files[0];
    if (droppedFile && (droppedFile.name.endsWith('.xlsx') || droppedFile.name.endsWith('.xls'))) {
      setFile(droppedFile);
      setError(null);
      setResult(null);
      setLogs([]);
    } else {
      setError('Vui lòng chọn file Excel (.xlsx hoặc .xls)');
    }
  };

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile) {
      setFile(selectedFile);
      setError(null);
      setResult(null);
      setLogs([]);
    }
  };

  const clearFile = () => {
    setFile(null);
    setResult(null);
    setError(null);
    setLogs([]);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleProcess = async () => {
    if (!file) return;
    if (targetRows <= 0) {
      setError('Số dòng phải lớn hơn 0');
      return;
    }

    setIsProcessing(true);
    setLogs([]);
    setError(null);
    
    try {
      const { buffer, stats } = await processExcel(file, parseInt(targetRows, 10), (msg) => {
        setLogs(prev => [...prev, msg]);
      });
      setResult({ buffer, stats });
    } catch (err) {
      setError(err.message || 'Có lỗi xảy ra trong quá trình xử lý');
      setLogs(prev => [...prev, `[LỖI] ${err.message}`]);
    } finally {
      setIsProcessing(false);
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
            onChange={handleFileChange} 
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
            {!isProcessing && (
              <button className="remove-file" onClick={clearFile} title="Xóa file">
                <X size={20} />
              </button>
            )}
          </div>

          {!result && !isProcessing && (
            <>
              <div className="form-group">
                <label>Tổng số dòng dữ liệu mong muốn (Bao gồm cả cũ và mới):</label>
                <input 
                  type="number" 
                  className="input-field"
                  value={targetRows}
                  onChange={(e) => setTargetRows(e.target.value)}
                  min="1"
                />
              </div>

              {error && (
                <div className="alert alert-error">
                  <AlertCircle size={20} />
                  <span>{error}</span>
                </div>
              )}

              <button 
                className="btn" 
                style={{ width: '100%', justifyContent: 'center', marginTop: '1rem' }}
                onClick={handleProcess}
              >
                <FilePlus2 size={20} />
                Bắt đầu Xử lý
              </button>
            </>
          )}

          {(isProcessing || logs.length > 0) && (
            <div className="logs-container" style={{ margin: '1.5rem 0', background: 'rgba(255,255,255,0.5)', borderRadius: '12px', padding: '1rem', border: '1px solid var(--card-border)', maxHeight: '250px', overflowY: 'auto' }}>
              <h4 style={{ marginBottom: '0.8rem', color: 'var(--text-main)', fontSize: '0.95rem' }}>Tiến trình xử lý:</h4>
              {logs.map((log, index) => (
                <div key={index} style={{ fontSize: '0.9rem', color: log.includes('[LỖI]') ? 'var(--error)' : 'var(--text-muted)', marginBottom: '0.5rem', display: 'flex', alignItems: 'flex-start', gap: '0.5rem' }}>
                  <ChevronRight size={16} style={{ flexShrink: 0, marginTop: '2px', color: log.includes('[LỖI]') ? 'var(--error)' : 'var(--primary)' }} />
                  <span>{log}</span>
                </div>
              ))}
              {isProcessing && (
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
