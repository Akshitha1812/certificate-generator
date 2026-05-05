import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { PDFDocument, rgb, StandardFonts } from 'pdf-lib';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';
import JSZip from 'jszip';
import { UploadCloud, FileSpreadsheet, FileImage, Download, CheckCircle, Settings, FileText, Award, LayoutGrid } from 'lucide-react';
import './index.css';

// Set up PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

function App() {
  const [appMode, setAppMode] = useState('certificate'); // 'certificate' or 'nametag'
  
  const [templateFile, setTemplateFile] = useState(null);
  const [templateType, setTemplateType] = useState(null); // 'image' or 'pdf'
  const [templateDimensions, setTemplateDimensions] = useState({ width: 0, height: 0 });
  const [excelFile, setExcelFile] = useState(null);
  const [participants, setParticipants] = useState([]); // [{name, role}]
  
  // Text positioning state
  const [nameConfig, setNameConfig] = useState({
    x: 100, y: 100, fontSize: 48, fontFamily: 'Helvetica', color: '#000000', align: 'center'
  });
  
  const [roleConfig, setRoleConfig] = useState({
    x: 100, y: 160, fontSize: 24, fontFamily: 'Helvetica', color: '#475569', align: 'center'
  });

  const [gridConfig, setGridConfig] = useState({
    gapX: 300, gapY: 200
  });

  const [isGenerating, setIsGenerating] = useState(false);
  const [outputMode, setOutputMode] = useState('separate');
  const [canvasScale, setCanvasScale] = useState(1);
  const [draggingElement, setDraggingElement] = useState(null); // 'name' or 'role'
  
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
      
      let nameColumn = '';
      let roleColumn = '';
      if (data.length > 0) {
        const keys = Object.keys(data[0]);
        nameColumn = keys.find(k => k.toLowerCase().includes('name')) || keys[0];
        roleColumn = keys.find(k => k.toLowerCase().includes('role') || k.toLowerCase().includes('title')) || '';
      }

      const extracted = data
        .map(row => ({
          name: row[nameColumn] ? String(row[nameColumn]).trim() : '',
          role: roleColumn && row[roleColumn] ? String(row[roleColumn]).trim() : ''
        }))
        .filter(p => p.name.length > 0);
      
      setParticipants(extracted);
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
        // Set initial positions
        setNameConfig(prev => ({ ...prev, x: img.width / 2, y: img.height / 3 }));
        setRoleConfig(prev => ({ ...prev, x: img.width / 2, y: img.height / 2 }));
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
      setNameConfig(prev => ({ ...prev, x: viewport.width / 2, y: viewport.height / 3 }));
      setRoleConfig(prev => ({ ...prev, x: viewport.width / 2, y: viewport.height / 2 }));
      
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

    if (containerRef.current) {
      const containerWidth = containerRef.current.clientWidth;
      const scale = containerWidth / canvas.width;
      setCanvasScale(scale < 1 ? scale : 1);
    }
  };

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

  const handleDrag = (e) => {
    if (e.buttons !== 1 || !draggingElement) return;
    
    const movementX = e.movementX / canvasScale;
    const movementY = e.movementY / canvasScale;

    if (draggingElement === 'name') {
      setNameConfig(prev => ({
        ...prev,
        x: Math.max(0, Math.min(prev.x + movementX, templateDimensions.width)),
        y: Math.max(0, Math.min(prev.y + movementY, templateDimensions.height))
      }));
    } else if (draggingElement === 'role') {
      setRoleConfig(prev => ({
        ...prev,
        x: Math.max(0, Math.min(prev.x + movementX, templateDimensions.width)),
        y: Math.max(0, Math.min(prev.y + movementY, templateDimensions.height))
      }));
    }
  };

  // Generate Final PDFs
  const generateCertificates = async () => {
    if (!templateFile || participants.length === 0) return;
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
      
      const nameColor = hexToRgb(nameConfig.color);
      const roleColor = hexToRgb(roleConfig.color);

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

      // Chunk participants based on mode
      let chunks = [];
      if (appMode === 'nametag') {
        for (let i = 0; i < participants.length; i += 6) {
          chunks.push(participants.slice(i, i + 6));
        }
      } else {
        chunks = participants.map(p => [p]);
      }

      for (let chunkIdx = 0; chunkIdx < chunks.length; chunkIdx++) {
        const chunk = chunks[chunkIdx];
        const docToEdit = outputMode === 'consolidated' ? consolidatedPdf : await PDFDocument.create();
        
        // Embed fonts
        const nameFont = await docToEdit.embedFont(
          nameConfig.fontFamily === 'Times-Roman' ? StandardFonts.TimesRoman :
          nameConfig.fontFamily === 'Courier' ? StandardFonts.Courier : StandardFonts.HelveticaBold
        );

        const roleFont = await docToEdit.embedFont(
          roleConfig.fontFamily === 'Times-Roman' ? StandardFonts.TimesRoman :
          roleConfig.fontFamily === 'Courier' ? StandardFonts.Courier : StandardFonts.HelveticaBold
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

        // Draw names/roles for this chunk
        for (let i = 0; i < chunk.length; i++) {
          const participant = chunk[i];
          const col = appMode === 'nametag' ? i % 2 : 0;
          const row = appMode === 'nametag' ? Math.floor(i / 2) : 0;

          // Draw Name
          const nameXOffset = col * gridConfig.gapX;
          const nameYOffset = row * gridConfig.gapY;
          
          const textWidth = nameFont.widthOfTextAtSize(participant.name, nameConfig.fontSize);
          let drawX = nameConfig.x + nameXOffset;
          if (nameConfig.align === 'center') drawX -= textWidth / 2;
          if (nameConfig.align === 'right') drawX -= textWidth;

          const pdfY = height - (nameConfig.y + nameYOffset) - (nameConfig.fontSize / 3);

          page.drawText(participant.name, {
            x: drawX, y: pdfY, size: nameConfig.fontSize, font: nameFont, color: rgb(nameColor.r, nameColor.g, nameColor.b)
          });

          // Draw Role (if in nametag mode and exists)
          if (appMode === 'nametag' && participant.role) {
            const roleXOffset = col * gridConfig.gapX;
            const roleYOffset = row * gridConfig.gapY;
            
            const roleTextWidth = roleFont.widthOfTextAtSize(participant.role, roleConfig.fontSize);
            let roleDrawX = roleConfig.x + roleXOffset;
            if (roleConfig.align === 'center') roleDrawX -= roleTextWidth / 2;
            if (roleConfig.align === 'right') roleDrawX -= roleTextWidth;

            const rolePdfY = height - (roleConfig.y + roleYOffset) - (roleConfig.fontSize / 3);

            page.drawText(participant.role, {
              x: roleDrawX, y: rolePdfY, size: roleConfig.fontSize, font: roleFont, color: rgb(roleColor.r, roleColor.g, roleColor.b)
            });
          }
        }

        if (outputMode === 'separate') {
          const pdfBytes = await docToEdit.save();
          // Use first name in chunk for separate file naming
          const safeName = chunk[0].name.replace(/[^a-z0-9]/gi, '_');
          const prefix = appMode === 'nametag' ? 'Nametags_' : 'Certificate_';
          zip.file(`${prefix}${safeName}.pdf`, pdfBytes);
        }
      }

      if (outputMode === 'consolidated') {
        const pdfBytes = await consolidatedPdf.save();
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `All_${appMode === 'nametag' ? 'Nametags' : 'Certificates'}_${new Date().toISOString().slice(0,10)}.pdf`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      } else {
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(zipBlob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${appMode === 'nametag' ? 'Nametags' : 'Certificates'}_ZIP_${new Date().toISOString().slice(0,10)}.zip`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }

    } catch (error) {
      console.error("Error generating documents:", error);
      alert("An error occurred while generating the documents.");
    } finally {
      setIsGenerating(false);
    }
  };

  // Render ghost overlays for nametag mode
  const renderGhostText = (config, isRole) => {
    if (appMode !== 'nametag') return null;
    const ghosts = [];
    for (let i = 1; i < 6; i++) { // 1 through 5
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = config.x + (col * gridConfig.gapX);
      const y = config.y + (row * gridConfig.gapY);
      
      ghosts.push(
        <div 
          key={`ghost-${isRole ? 'role' : 'name'}-${i}`}
          style={{
            position: 'absolute',
            left: `${x * canvasScale}px`,
            top: `${y * canvasScale}px`,
            transform: config.align === 'center' ? 'translate(-50%, -50%)' : config.align === 'right' ? 'translate(-100%, -50%)' : 'translate(0, -50%)',
            fontSize: `${config.fontSize * canvasScale}px`,
            fontFamily: config.fontFamily,
            color: config.color,
            fontWeight: 'bold',
            opacity: 0.3,
            pointerEvents: 'none',
            whiteSpace: 'nowrap'
          }}
        >
          {isRole ? 'Sample Role' : 'Sample Name'}
        </div>
      );
    }
    return ghosts;
  };

  return (
    <div className="app-container">
      <div className="header">
        <h1>Document Generator</h1>
        
        <div style={{ display: 'flex', justifyContent: 'center', gap: '1rem', marginTop: '1.5rem' }}>
          <button 
            className={`btn ${appMode === 'certificate' ? '' : 'inactive'}`}
            style={{ opacity: appMode === 'certificate' ? 1 : 0.5, background: appMode === 'certificate' ? 'var(--accent)' : 'transparent', border: '1px solid var(--accent)' }}
            onClick={() => setAppMode('certificate')}
          >
            <Award size={20} /> Certificate Mode
          </button>
          <button 
            className={`btn ${appMode === 'nametag' ? '' : 'inactive'}`}
            style={{ opacity: appMode === 'nametag' ? 1 : 0.5, background: appMode === 'nametag' ? 'var(--accent)' : 'transparent', border: '1px solid var(--accent)' }}
            onClick={() => setAppMode('nametag')}
          >
            <LayoutGrid size={20} /> Nametag Mode (6/Page)
          </button>
        </div>
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
                <p>{participants.length} names found</p>
              </>
            ) : (
              <>
                <FileSpreadsheet size={48} />
                <h3>Upload Excel Data</h3>
                <p>File must contain a "Name" column {appMode === 'nametag' ? 'and optional "Role" column' : ''}</p>
              </>
            )}
          </div>
        </div>

        {/* Configuration Section */}
        {templateFile && (
          <div className="preview-section">
            <div className="controls-panel">
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginRight: '1rem', width: '100%' }}>
                <Settings size={20} color="var(--accent)" />
                <h3 style={{ margin: 0 }}>Name Text Settings</h3>
              </div>
              <div className="control-group">
                <label>Font Size</label>
                <input type="number" value={nameConfig.fontSize} onChange={(e) => setNameConfig({...nameConfig, fontSize: Number(e.target.value)})} style={{ width: '80px' }} />
              </div>
              <div className="control-group">
                <label>Color</label>
                <input type="color" value={nameConfig.color} onChange={(e) => setNameConfig({...nameConfig, color: e.target.value})} />
              </div>
            </div>

            {appMode === 'nametag' && (
              <div className="controls-panel">
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginRight: '1rem', width: '100%' }}>
                  <Settings size={20} color="var(--success)" />
                  <h3 style={{ margin: 0 }}>Role Text Settings</h3>
                </div>
                <div className="control-group">
                  <label>Font Size</label>
                  <input type="number" value={roleConfig.fontSize} onChange={(e) => setRoleConfig({...roleConfig, fontSize: Number(e.target.value)})} style={{ width: '80px' }} />
                </div>
                <div className="control-group">
                  <label>Color</label>
                  <input type="color" value={roleConfig.color} onChange={(e) => setRoleConfig({...roleConfig, color: e.target.value})} />
                </div>
              </div>
            )}

            {appMode === 'nametag' && (
              <div className="controls-panel" style={{ background: 'rgba(59, 130, 246, 0.1)' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginRight: '1rem', width: '100%' }}>
                  <LayoutGrid size={20} color="var(--accent)" />
                  <h3 style={{ margin: 0 }}>Grid Alignment (2 Columns x 3 Rows)</h3>
                </div>
                <div className="control-group" style={{ width: '45%' }}>
                  <label>Horizontal Spacing (X)</label>
                  <input type="range" min="0" max={templateDimensions.width} value={gridConfig.gapX} onChange={(e) => setGridConfig({...gridConfig, gapX: Number(e.target.value)})} />
                </div>
                <div className="control-group" style={{ width: '45%' }}>
                  <label>Vertical Spacing (Y)</label>
                  <input type="range" min="0" max={templateDimensions.height} value={gridConfig.gapY} onChange={(e) => setGridConfig({...gridConfig, gapY: Number(e.target.value)})} />
                </div>
              </div>
            )}
            
            <p style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>
              Drag the sample text below to position it perfectly on the top-left section of your template.
            </p>

            <div className="canvas-container" ref={containerRef} onMouseUp={handleDragEnd} onMouseLeave={handleDragEnd}>
              <canvas ref={canvasRef} />
              
              {/* Draggable Overlay: Name */}
              <div 
                className="overlay-text-preview"
                style={{
                  position: 'absolute',
                  left: `${nameConfig.x * canvasScale}px`,
                  top: `${nameConfig.y * canvasScale}px`,
                  transform: nameConfig.align === 'center' ? 'translate(-50%, -50%)' : nameConfig.align === 'right' ? 'translate(-100%, -50%)' : 'translate(0, -50%)',
                  fontSize: `${nameConfig.fontSize * canvasScale}px`,
                  fontFamily: nameConfig.fontFamily,
                  color: nameConfig.color,
                  fontWeight: 'bold',
                  cursor: 'move',
                  userSelect: 'none',
                  textShadow: '0px 0px 2px rgba(255,255,255,0.8)',
                  border: '1px dashed rgba(59, 130, 246, 0.5)',
                  padding: '2px 8px',
                  whiteSpace: 'nowrap',
                  zIndex: 10
                }}
                onMouseDown={(e) => handleDragStart(e, 'name')}
                onMouseMove={handleDrag}
              >
                Sample Name
              </div>

              {/* Draggable Overlay: Role (Nametag Mode only) */}
              {appMode === 'nametag' && (
                <div 
                  className="overlay-text-preview"
                  style={{
                    position: 'absolute',
                    left: `${roleConfig.x * canvasScale}px`,
                    top: `${roleConfig.y * canvasScale}px`,
                    transform: roleConfig.align === 'center' ? 'translate(-50%, -50%)' : roleConfig.align === 'right' ? 'translate(-100%, -50%)' : 'translate(0, -50%)',
                    fontSize: `${roleConfig.fontSize * canvasScale}px`,
                    fontFamily: roleConfig.fontFamily,
                    color: roleConfig.color,
                    fontWeight: 'bold',
                    cursor: 'move',
                    userSelect: 'none',
                    textShadow: '0px 0px 2px rgba(255,255,255,0.8)',
                    border: '1px dashed rgba(16, 185, 129, 0.5)',
                    padding: '2px 8px',
                    whiteSpace: 'nowrap',
                    zIndex: 10
                  }}
                  onMouseDown={(e) => handleDragStart(e, 'role')}
                  onMouseMove={handleDrag}
                >
                  Sample Role
                </div>
              )}

              {/* Ghost Grid Previews */}
              {renderGhostText(nameConfig, false)}
              {renderGhostText(roleConfig, true)}

            </div>
          </div>
        )}

        {/* Names Preview Section */}
        {participants.length > 0 && (
          <div className="names-preview">
            <h4 style={{ marginBottom: '1rem', color: 'var(--text-secondary)' }}>Extracted Participants ({participants.length})</h4>
            <ul>
              {participants.slice(0, 20).map((p, i) => (
                <li key={i}><CheckCircle /> {p.name} {p.role && <span style={{opacity: 0.6}}>- {p.role}</span>}</li>
              ))}
              {participants.length > 20 && <li>...and {participants.length - 20} more</li>}
            </ul>
          </div>
        )}

        {/* Generate Button */}
        {templateFile && excelFile && participants.length > 0 && (
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
                <><Download size={24} /> Generate {participants.length} {appMode === 'nametag' ? 'Nametags' : 'Certificates'}</>
              )}
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;
