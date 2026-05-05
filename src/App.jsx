import React, { useState, useRef, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { PDFDocument, rgb, StandardFonts } from 'pdf-lib';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';
import JSZip from 'jszip';
import { UploadCloud, FileSpreadsheet, FileImage, Download, CheckCircle, Settings, FileText } from 'lucide-react';
import './index.css';

// Set up PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

function App() {
  const [templateFile, setTemplateFile] = useState(null);
  const [templateType, setTemplateType] = useState(null); // 'image' or 'pdf'
  const [templateDimensions, setTemplateDimensions] = useState({ width: 0, height: 0 });
  const [excelFile, setExcelFile] = useState(null);
  const [names, setNames] = useState([]);
  
  // Text positioning state (relative to actual image/pdf dimensions)
  const [textConfig, setTextConfig] = useState({
    x: 100,
    y: 100,
    fontSize: 48,
    fontFamily: 'Helvetica',
    color: '#000000',
    align: 'center' // center, left, right
  });

  const [isGenerating, setIsGenerating] = useState(false);
  const [outputMode, setOutputMode] = useState('separate');
  const [canvasScale, setCanvasScale] = useState(1);
  
  const canvasRef = useRef(null);
  const containerRef = useRef(null);

  // Parse Excel File
  const handleExcelUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setExcelFile(file);

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);
      
      // Look for a column named 'Name' (case-insensitive)
      let nameColumn = '';
      if (data.length > 0) {
        const keys = Object.keys(data[0]);
        nameColumn = keys.find(k => k.toLowerCase() === 'name') || keys[0]; // fallback to first column
      }

      const extractedNames = data
        .map(row => row[nameColumn])
        .filter(name => name && typeof name === 'string' && name.trim().length > 0);
      
      setNames(extractedNames);
    };
    reader.readAsBinaryString(file);
  };

  // Handle Template Upload
  const handleTemplateUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setTemplateFile(file);
    
    if (file.type.includes('pdf')) {
      setTemplateType('pdf');
      renderPdfPreview(file);
    } else if (file.type.includes('image')) {
      setTemplateType('image');
      renderImagePreview(file);
    } else {
      alert("Please upload a valid PDF or Image file.");
    }
  };

  const renderImagePreview = (file) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      const img = new Image();
      img.onload = () => {
        setTemplateDimensions({ width: img.width, height: img.height });
        // Set initial text position to center
        setTextConfig(prev => ({
          ...prev,
          x: img.width / 2,
          y: img.height / 2
        }));
        drawCanvas(img, null);
      };
      img.src = event.target.result;
    };
    reader.readAsDataURL(file);
  };

  const renderPdfPreview = async (file) => {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      const page = await pdf.getPage(1);
      const viewport = page.getViewport({ scale: 1.0 }); // Use 1.0 so coordinates exactly match PDF points

      setTemplateDimensions({ width: viewport.width, height: viewport.height });
      // Set initial text position to center
      setTextConfig(prev => ({
        ...prev,
        x: viewport.width / 2,
        y: viewport.height / 2
      }));
      
      drawCanvas(null, page, viewport);
    } catch (err) {
      console.error("Error rendering PDF preview:", err);
      alert("Failed to render PDF preview. Please check the console for details.");
    }
  };

  const drawCanvas = async (imgSource, pdfPage, viewport = null) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    if (imgSource) {
      canvas.width = imgSource.width;
      canvas.height = imgSource.height;
      ctx.drawImage(imgSource, 0, 0);
    } else if (pdfPage) {
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      
      const renderContext = {
        canvasContext: ctx,
        viewport: viewport
      };
      await pdfPage.render(renderContext).promise;
    }

    // Calculate scale factor for overlay rendering
    if (containerRef.current) {
      const containerWidth = containerRef.current.clientWidth;
      const scale = containerWidth / canvas.width;
      setCanvasScale(scale < 1 ? scale : 1);
    }
  };

  // Re-draw when file changes (handled in specific render functions to avoid infinite loops)
  // But we need to handle resize to update scale
  useEffect(() => {
    const handleResize = () => {
      if (canvasRef.current && containerRef.current) {
        const scale = containerRef.current.clientWidth / canvasRef.current.width;
        setCanvasScale(scale < 1 ? scale : 1);
      }
    };
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  // Handle Dragging the Text Box
  const handleDrag = (e) => {
    if (e.buttons !== 1) return; // Only drag on left click
    
    // Calculate new position based on movement and scale
    const movementX = e.movementX / canvasScale;
    const movementY = e.movementY / canvasScale;

    setTextConfig(prev => ({
      ...prev,
      x: Math.max(0, Math.min(prev.x + movementX, templateDimensions.width)),
      y: Math.max(0, Math.min(prev.y + movementY, templateDimensions.height))
    }));
  };

  // Generate Final PDFs
  const generateCertificates = async () => {
    if (!templateFile || names.length === 0) return;
    setIsGenerating(true);

    try {
      const hexToRgb = (hex) => {
        const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
        return result ? {
          r: parseInt(result[1], 16) / 255,
          g: parseInt(result[2], 16) / 255,
          b: parseInt(result[3], 16) / 255
        } : { r: 0, g: 0, b: 0 };
      };
      
      const textColor = hexToRgb(textConfig.color);

      // Load Template Source once
      const templateBuffer = await templateFile.arrayBuffer();
      let templatePdfDoc = null;
      let templateImageBytes = templateBuffer;
      let imageType = 'jpg';
      
      if (templateType === 'pdf') {
         templatePdfDoc = await PDFDocument.load(templateBuffer);
      } else if (templateFile.type.includes('png')) {
         imageType = 'png';
      }

      let consolidatedPdf;
      let zip;
      
      if (outputMode === 'consolidated') {
        consolidatedPdf = await PDFDocument.create();
      } else {
        zip = new JSZip();
      }

      for (let i = 0; i < names.length; i++) {
        const name = names[i];
        const docToEdit = outputMode === 'consolidated' ? consolidatedPdf : await PDFDocument.create();
        
        // Embed font
        const font = await docToEdit.embedFont(
          textConfig.fontFamily === 'Times-Roman' ? StandardFonts.TimesRoman :
          textConfig.fontFamily === 'Courier' ? StandardFonts.Courier :
          StandardFonts.HelveticaBold
        );

        let page;
        let width, height;

        if (templateType === 'pdf') {
          const [copiedPage] = await docToEdit.copyPages(templatePdfDoc, [0]);
          page = docToEdit.addPage(copiedPage);
          width = page.getWidth();
          height = page.getHeight();
        } else {
          let templateImageObj;
          if (imageType === 'png') {
            templateImageObj = await docToEdit.embedPng(templateImageBytes);
          } else {
            templateImageObj = await docToEdit.embedJpg(templateImageBytes);
          }
          const dims = templateImageObj.scale(1);
          width = dims.width;
          height = dims.height;
          page = docToEdit.addPage([width, height]);
          page.drawImage(templateImageObj, { x: 0, y: 0, width, height });
        }

        const textWidth = font.widthOfTextAtSize(name, textConfig.fontSize);
        let drawX = textConfig.x;
        if (textConfig.align === 'center') drawX -= textWidth / 2;
        if (textConfig.align === 'right') drawX -= textWidth;

        const pdfY = height - textConfig.y - (textConfig.fontSize / 3);

        page.drawText(name, {
          x: drawX,
          y: pdfY,
          size: textConfig.fontSize,
          font: font,
          color: rgb(textColor.r, textColor.g, textColor.b),
        });

        if (outputMode === 'separate') {
          const pdfBytes = await docToEdit.save();
          const safeName = name.replace(/[^a-z0-9]/gi, '_');
          zip.file(`${safeName}_certificate.pdf`, pdfBytes);
        }
      }

      if (outputMode === 'consolidated') {
        const pdfBytes = await consolidatedPdf.save();
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `All_Certificates_${new Date().toISOString().slice(0,10)}.pdf`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      } else {
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(zipBlob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `Certificates_ZIP_${new Date().toISOString().slice(0,10)}.zip`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }

    } catch (error) {
      console.error("Error generating certificates:", error);
      alert("An error occurred while generating the certificates.");
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="app-container">
      <div className="header">
        <h1>Certificate Generator</h1>
        <p>Upload your template and an Excel file to generate personalized certificates instantly.</p>
      </div>

      <div className="glass-card">
        <div className="upload-grid">
          {/* Template Upload */}
          <div className={`dropzone ${templateFile ? 'active' : ''}`} onClick={() => document.getElementById('template-upload').click()}>
            <input 
              id="template-upload" 
              type="file" 
              accept=".pdf,.jpg,.jpeg,.png" 
              style={{ display: 'none' }} 
              onChange={handleTemplateUpload}
            />
            {templateFile ? (
              <>
                {templateType === 'pdf' ? <FileText size={48} /> : <FileImage size={48} />}
                <h3>{templateFile.name}</h3>
                <p>Template loaded successfully</p>
              </>
            ) : (
              <>
                <UploadCloud size={48} />
                <h3>Upload Template</h3>
                <p>Drag & drop or click to browse (PDF, JPG, PNG)</p>
              </>
            )}
          </div>

          {/* Excel Upload */}
          <div className={`dropzone ${excelFile ? 'active' : ''}`} onClick={() => document.getElementById('excel-upload').click()}>
            <input 
              id="excel-upload" 
              type="file" 
              accept=".xlsx,.xls,.csv" 
              style={{ display: 'none' }} 
              onChange={handleExcelUpload}
            />
            {excelFile ? (
              <>
                <FileSpreadsheet size={48} />
                <h3>{excelFile.name}</h3>
                <p>{names.length} names found</p>
              </>
            ) : (
              <>
                <FileSpreadsheet size={48} />
                <h3>Upload Names</h3>
                <p>Excel file with a "Name" column (.xlsx, .csv)</p>
              </>
            )}
          </div>
        </div>

        {/* Configuration Section (Visible only when template is loaded) */}
        {templateFile && (
          <div className="preview-section">
            <div className="controls-panel">
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginRight: '1rem' }}>
                <Settings size={20} color="var(--accent)" />
                <h3 style={{ margin: 0 }}>Text Settings</h3>
              </div>
              <div className="control-group">
                <label>Font Size</label>
                <input 
                  type="number" 
                  value={textConfig.fontSize} 
                  onChange={(e) => setTextConfig({...textConfig, fontSize: Number(e.target.value)})}
                  style={{ width: '80px' }}
                />
              </div>
              <div className="control-group">
                <label>Color</label>
                <input 
                  type="color" 
                  value={textConfig.color} 
                  onChange={(e) => setTextConfig({...textConfig, color: e.target.value})}
                />
              </div>
              <div className="control-group">
                <label>Font</label>
                <select 
                  value={textConfig.fontFamily}
                  onChange={(e) => setTextConfig({...textConfig, fontFamily: e.target.value})}
                >
                  <option value="Helvetica">Helvetica (Bold)</option>
                  <option value="Times-Roman">Times New Roman</option>
                  <option value="Courier">Courier</option>
                </select>
              </div>
              <div className="control-group">
                <label>Alignment</label>
                <select 
                  value={textConfig.align}
                  onChange={(e) => setTextConfig({...textConfig, align: e.target.value})}
                >
                  <option value="center">Center</option>
                  <option value="left">Left</option>
                  <option value="right">Right</option>
                </select>
              </div>
            </div>
            
            <p style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>
              Drag the sample text below to position it exactly where the names should appear.
            </p>

            <div className="canvas-container" ref={containerRef}>
              <canvas ref={canvasRef} />
              
              {/* Draggable Overlay */}
              <div 
                className="overlay-text-preview"
                style={{
                  position: 'absolute',
                  left: `${textConfig.x * canvasScale}px`,
                  top: `${textConfig.y * canvasScale}px`,
                  transform: textConfig.align === 'center' ? 'translate(-50%, -50%)' : textConfig.align === 'right' ? 'translate(-100%, -50%)' : 'translate(0, -50%)',
                  fontSize: `${textConfig.fontSize * canvasScale}px`,
                  fontFamily: textConfig.fontFamily,
                  color: textConfig.color,
                  fontWeight: 'bold',
                  cursor: 'move',
                  userSelect: 'none',
                  textShadow: '0px 0px 2px rgba(255,255,255,0.8)',
                  border: '1px dashed rgba(59, 130, 246, 0.5)',
                  padding: '2px 8px',
                  whiteSpace: 'nowrap'
                }}
                onMouseMove={handleDrag}
              >
                Sample Name
              </div>
            </div>
          </div>
        )}

        {/* Names Preview Section */}
        {names.length > 0 && (
          <div className="names-preview">
            <h4 style={{ marginBottom: '1rem', color: 'var(--text-secondary)' }}>Extracted Names ({names.length})</h4>
            <ul>
              {names.slice(0, 20).map((name, i) => (
                <li key={i}><CheckCircle /> {name}</li>
              ))}
              {names.length > 20 && <li>...and {names.length - 20} more</li>}
            </ul>
          </div>
        )}

        {/* Generate Button */}
        {templateFile && excelFile && names.length > 0 && (
          <div className="generate-section">
            <div style={{ marginBottom: '1.5rem', display: 'flex', justifyContent: 'center', gap: '2rem', color: 'var(--text-secondary)' }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', cursor: 'pointer' }}>
                <input 
                  type="radio" 
                  name="outputMode" 
                  value="separate" 
                  checked={outputMode === 'separate'} 
                  onChange={() => setOutputMode('separate')} 
                  style={{ cursor: 'pointer' }}
                />
                Separate PDFs (ZIP)
              </label>
              <label style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', cursor: 'pointer' }}>
                <input 
                  type="radio" 
                  name="outputMode" 
                  value="consolidated" 
                  checked={outputMode === 'consolidated'} 
                  onChange={() => setOutputMode('consolidated')} 
                  style={{ cursor: 'pointer' }}
                />
                Single Consolidated PDF
              </label>
            </div>
            <button 
              className="btn" 
              onClick={generateCertificates}
              disabled={isGenerating}
              style={{ fontSize: '1.2rem', padding: '1rem 3rem' }}
            >
              {isGenerating ? (
                <><span className="loading-spinner"></span> Generating PDFs...</>
              ) : (
                <><Download size={24} /> Generate {names.length} Certificates</>
              )}
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;
