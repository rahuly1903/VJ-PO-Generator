import React, { useState } from 'react';
import { Upload, Download, CheckCircle, AlertCircle, Loader, FileText, Package, X } from 'lucide-react';
import logo from './assets/renaissance-logo.webp';

const formatSize = (bytes) => {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
};

const POGenerator = () => {
  const [file, setFile] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [status, setStatus] = useState({ type: '', message: '' });
  const [generatedFiles, setGeneratedFiles] = useState([]);
  const [librariesLoaded, setLibrariesLoaded] = useState(false);
  const [dragActive, setDragActive] = useState(false);

  // Load external libraries
  React.useEffect(() => {
    const loadLibraries = async () => {
      // Load XLSX
      if (!window.XLSX) {
        const xlsxScript = document.createElement('script');
        xlsxScript.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
        document.head.appendChild(xlsxScript);
        await new Promise(resolve => xlsxScript.onload = resolve);
      }
      
      // Load JSZip
      if (!window.JSZip) {
        const jszipScript = document.createElement('script');
        jszipScript.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
        document.head.appendChild(jszipScript);
        await new Promise(resolve => jszipScript.onload = resolve);
      }
      
      setLibrariesLoaded(true);
    };
    
    loadLibraries();
  }, []);

  const assignFile = (selectedFile) => {
    if (!selectedFile) return;
    const name = selectedFile.name.toLowerCase();
    if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
      setFile(selectedFile);
      setGeneratedFiles([]);
      setStatus({ type: '', message: '' });
    } else {
      setStatus({ type: 'error', message: 'Use an Excel file (.xlsx or .xls).' });
      setFile(null);
    }
  };

  const handleFileChange = (e) => {
    assignFile(e.target.files?.[0]);
    e.target.value = '';
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setDragActive(false);
    if (processing) return;
    assignFile(e.dataTransfer.files?.[0]);
  };

  const clearFile = () => {
    setFile(null);
    setGeneratedFiles([]);
    setStatus({ type: '', message: '' });
  };

  const parseExcel = async (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = window.XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = window.XLSX.utils.sheet_to_json(worksheet);
          resolve(jsonData);
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  };

  const escapeXml = (text) => {
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  };

  const generateDocxXml = (row) => {
    const orderNo = String(row["Order#"] || "").replace("#", "").trim();
    const customer = String(row["CustomerName"] || "").trim().replace("/", "-");
    const fmSku = String(row["FM Setting SKU"] || "").trim();
    const manufacturerSku = String(row["ManufacturerSKU"] || "").trim();
    const productName = String(row["Product Name"] || "").trim();
    const orderDate = String(row["OrderDate"] || "").trim();

    const addField = (label, value) => {
      if (value && String(value).trim()) {
        return `
          <w:p>
            <w:r><w:rPr><w:b/></w:rPr><w:t>${escapeXml(label)}: </w:t></w:r>
            <w:r><w:t>${escapeXml(value)}</w:t></w:r>
          </w:p>`;
      }
      return '';
    };

    const title = `#${orderNo} - ${manufacturerSku} - ${customer} - ${productName} / Malakan Diamond Co./LINE ITEM - 1`;

    let threeStoneSection = '';
    if (productName.toLowerCase().includes("three stone")) {
      threeStoneSection = `
        <w:p>
          <w:r><w:rPr><w:b/></w:rPr><w:t>Three Stone (if Applicable): </w:t></w:r>
          <w:r><w:t>Y</w:t></w:r>
        </w:p>
        <w:p>
          <w:r><w:rPr><w:b/></w:rPr><w:t>Side Stone Shape: </w:t></w:r>
          <w:r><w:t></w:t></w:r>
        </w:p>`;
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr><w:b/><w:sz w:val="24"/></w:rPr>
        <w:t>${escapeXml(title)}</w:t>
      </w:r>
    </w:p>
    <w:p/>
    
    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Order Information</w:t></w:r></w:p>
    ${addField("Order #", orderNo)}
    ${addField("Date Created", orderDate)}
    ${addField("Manufacture", "Malakan Diamond Co.")}
    ${addField("JD SKU", fmSku)}
    ${addField("Manufacturer SKU", manufacturerSku)}
    ${addField("Product Name", productName)}
    <w:p/>
    
    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Client Details</w:t></w:r></w:p>
    ${addField("Client Name", row["CustomerName"])}
    ${addField("Band Style", row["POLISH"])}
    ${addField("Diamond Shape", row["Gemstone Shape"])}
    ${addField("Diamond Carat Weight", row["Gemstone Ct"])}
    ${addField("Diamond Color", row["Dia Color"])}
    ${addField("Diamond Clarity", row["Dia Clarity"])}
    ${addField("Color Clarity", row["ColorClarity"])}
    ${addField("Diamond Vendor", row["DiamondVendorCode"])}
    ${addField("Diamond SKU", row["DiamondSKU"])}
    ${addField("GIA Certificate No", row["Certificate No"])}
    ${addField("Engraving", row["Engraving"])}
    ${addField("Metal", row["Setting Metal"])}
    ${addField("Ring Size", row["Setting Size"])}
    ${addField("Width", row["AVGWidth"])}
    ${addField("Signature Stone", row["SignatureStone"])}
    ${threeStoneSection}
    ${addField("Quantity", "1")}
    ${addField("Cost", row["VendorCost"])}
    ${addField("Remark", row["REMARK"])}
    
    <w:p/>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr><w:i/><w:sz w:val="18"/></w:rPr>
        <w:t>Generated Automatically - Purchase Order System</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;
  };

  const createDocxBlob = (xmlContent) => {
    const zip = new window.JSZip();

    // Add required files for a valid DOCX
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

    zip.folder('_rels').file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

    zip.folder('word').file('document.xml', xmlContent);

    zip.folder('word/_rels').file('document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`);

    return zip.generateAsync({ type: 'blob' });
  };

  const downloadBlob = (blob, filename) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const processExcel = async () => {
    if (!file) {
      setStatus({ type: 'error', message: 'Select a file first.' });
      return;
    }

    if (!librariesLoaded) {
      setStatus({ type: 'error', message: 'Libraries are still loading. Try again in a moment.' });
      return;
    }

    setProcessing(true);
    setGeneratedFiles([]);
    setStatus({ type: 'info', message: 'Reading the workbook…' });

    try {
      const jsonData = await parseExcel(file);

      if (jsonData.length === 0) {
        setStatus({ type: 'error', message: 'No rows found in the first sheet.' });
        setProcessing(false);
        return;
      }

      setStatus({ type: 'info', message: `Generating ${jsonData.length} purchase orders…` });

      const files = [];
      for (let i = 0; i < jsonData.length; i++) {
        const row = jsonData[i];
        const orderNo = String(row["Order#"] || "").replace("#", "").trim();
        const customer = String(row["CustomerName"] || "").trim().replace("/", "-");
        
        const xmlContent = generateDocxXml(row);
        const blob = await createDocxBlob(xmlContent);
        const filename = `PO_${orderNo}_${customer}.docx`.replace(/ /g, "_");
        
        files.push({ filename, blob, orderNo, customer });
      }

      setGeneratedFiles(files);
      setStatus({ 
        type: 'success', 
        message: `Generated ${files.length} purchase ${files.length === 1 ? 'order' : 'orders'}.` 
      });
    } catch (error) {
      console.error(error);
      setStatus({ 
        type: 'error', 
        message: `${error.message}. Check that the workbook matches the expected columns.` 
      });
    } finally {
      setProcessing(false);
    }
  };

  const downloadIndividual = (fileData) => {
    downloadBlob(fileData.blob, fileData.filename);
  };

  const downloadAll = async () => {
    setStatus({ type: 'info', message: 'Downloading files…' });
    
    for (let i = 0; i < generatedFiles.length; i++) {
      downloadBlob(generatedFiles[i].blob, generatedFiles[i].filename);
      await new Promise(resolve => setTimeout(resolve, 300));
    }
    
    setStatus({ type: 'success', message: 'All files downloaded.' });
  };

  const downloadAllAsZip = async () => {
    setStatus({ type: 'info', message: 'Preparing the ZIP…' });
    
    const zip = new window.JSZip();
    const folder = zip.folder('Purchase_Orders');
    
    for (const fileData of generatedFiles) {
      folder.file(fileData.filename, fileData.blob);
    }
    
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    downloadBlob(zipBlob, 'Purchase_Orders.zip');
    
    setStatus({ type: 'success', message: 'ZIP downloaded.' });
  };

  const ready = Boolean(file) && !processing && librariesLoaded;

  return (
    <div className="po">
      <header className="po-header">
        <img className="po-logo" src={logo} alt="Renaissance Global Limited" />
        <p className="po-kicker">Documents</p>
        <h1>Purchase orders</h1>
        <p className="po-lede">
          Upload an Excel workbook. Each row becomes a Word file you can download.
        </p>
      </header>

      <div className="po-layout">
        <section className="po-panel" aria-labelledby="upload-heading">
          <div className="po-panel-head">
            <h2 id="upload-heading">Source file</h2>
            <p>Excel, first sheet only.</p>
          </div>

          <div
            className={`po-drop ${dragActive ? 'is-active' : ''} ${file ? 'has-file' : ''}`}
            onDragOver={(e) => {
              e.preventDefault();
              if (!processing) setDragActive(true);
            }}
            onDragLeave={(e) => {
              if (!e.currentTarget.contains(e.relatedTarget)) setDragActive(false);
            }}
            onDrop={handleDrop}
          >
            <input
              type="file"
              accept=".xlsx,.xls,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              onChange={handleFileChange}
              id="file-upload"
              disabled={processing}
            />
            <label htmlFor="file-upload">
              <Upload size={18} strokeWidth={1.75} />
              <span>{dragActive ? 'Drop the file' : 'Choose a file or drop it here'}</span>
              <small>.xlsx or .xls</small>
            </label>
          </div>

          {file && (
            <div className="po-selected">
              <FileText size={16} strokeWidth={1.75} />
              <div>
                <strong>{file.name}</strong>
                <span>{formatSize(file.size)}</span>
              </div>
              <button type="button" onClick={clearFile} disabled={processing} aria-label="Remove file">
                <X size={16} strokeWidth={1.75} />
              </button>
            </div>
          )}

          <button
            type="button"
            className="po-primary"
            onClick={processExcel}
            disabled={!ready}
          >
            {!librariesLoaded || processing ? (
              <Loader className="po-spin" size={16} strokeWidth={1.75} />
            ) : null}
            {!librariesLoaded ? 'Loading' : processing ? 'Generating' : 'Generate'}
          </button>

          {status.message && (
            <div className={`po-status is-${status.type || 'info'}`} role="status">
              {status.type === 'success' ? (
                <CheckCircle size={16} strokeWidth={1.75} />
              ) : (
                <AlertCircle size={16} strokeWidth={1.75} />
              )}
              <p>{status.message}</p>
            </div>
          )}

          <div className="po-steps">
            <p>Instructions</p>
            <ol>
              <li>Upload your Excel file</li>
              <li>Click Generate</li>
              <li>Choose a download option</li>
            </ol>
          </div>
        </section>

        <section className="po-panel po-results" aria-labelledby="results-heading">
          <div className="po-results-head">
            <div className="po-panel-head">
              <h2 id="results-heading">
                Orders
                <span>{generatedFiles.length}</span>
              </h2>
              <p>{generatedFiles.length ? 'Download one file, every file, or a ZIP.' : 'Generated files appear here.'}</p>
            </div>
            {generatedFiles.length > 0 && (
              <div className="po-actions">
                <button type="button" onClick={downloadAll}>
                  <Download size={15} strokeWidth={1.75} />
                  All
                </button>
                <button type="button" onClick={downloadAllAsZip}>
                  <Package size={15} strokeWidth={1.75} />
                  ZIP
                </button>
              </div>
            )}
          </div>

          {generatedFiles.length === 0 ? (
            <div className="po-empty">
              <p>Nothing generated yet.</p>
            </div>
          ) : (
            <ul className="po-list">
              {generatedFiles.map((fileData, index) => (
                <li key={`${fileData.filename}-${index}`}>
                  <div>
                    <strong>{fileData.filename}</strong>
                    <span>
                      {fileData.orderNo ? `#${fileData.orderNo}` : 'No order number'}
                      {fileData.customer ? ` · ${fileData.customer}` : ''}
                    </span>
                  </div>
                  <button type="button" onClick={() => downloadIndividual(fileData)}>
                    Download
                  </button>
                </li>
              ))}
            </ul>
          )}
        </section>
      </div>
    </div>
  );
};

export default POGenerator;