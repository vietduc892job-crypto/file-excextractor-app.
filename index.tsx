/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useEffect, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { GoogleGenAI, Type } from "@google/genai";

// Make XLSX available from the global scope (loaded via CDN)
declare const XLSX: any;

const ApiKeyModal: React.FC<{ onApiKeySubmit: (key: string) => void }> = ({ onApiKeySubmit }) => {
    const [localApiKey, setLocalApiKey] = useState('');

    const handleSubmit = (e: React.FormEvent) => {
        e.preventDefault();
        if (localApiKey.trim()) {
            onApiKeySubmit(localApiKey.trim());
        }
    };

    return (
        <div className="modal-overlay">
            <div className="modal-content">
                <h2>Enter Your Google AI API Key</h2>
                <p>To use Data Alchemist, please provide your Google AI API key. Your key is not stored and is only used for this session.</p>
                <form onSubmit={handleSubmit}>
                    <input
                        type="password"
                        value={localApiKey}
                        onChange={(e) => setLocalApiKey(e.target.value)}
                        placeholder="Enter your API key here"
                        className="modal-input"
                        aria-label="Google AI API Key"
                    />
                    <button type="submit" className="btn btn-primary modal-button">Save Key & Start</button>
                </form>
                <a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noopener noreferrer" className="modal-link">Get your API key from Google AI Studio</a>
            </div>
        </div>
    );
};


const App: React.FC = () => {
    const [apiKey, setApiKey] = useState<string>('');
    const [isModalOpen, setIsModalOpen] = useState(true);
    const [uploadedFile, setUploadedFile] = useState<File | null>(null);
    const [fileType, setFileType] = useState<'image' | 'excel' | 'word' | 'pdf' | null>(null);
    const [previewUrl, setPreviewUrl] = useState<string | null>(null);
    const [extractedData, setExtractedData] = useState<{[sheetName: string]: string[][]}>({});
    const [extractedText, setExtractedText] = useState<string | null>(null);
    const [extractionMode, setExtractionMode] = useState<'excel' | 'word' | null>(null);
    const [activeSheetName, setActiveSheetName] = useState<string | null>(null);
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const [isDragging, setIsDragging] = useState(false);
    const [isSwordUnsheathed, setIsSwordUnsheathed] = useState(false);
    const [translateTo, setTranslateTo] = useState('none');
    const fileInputRef = useRef<HTMLInputElement>(null);

    const handleApiKeySubmit = (key: string) => {
        setApiKey(key);
        setIsModalOpen(false);
    };

    const resetState = useCallback(() => {
        setUploadedFile(null);
        setFileType(null);
        setPreviewUrl(null);
        setExtractedData({});
        setExtractedText(null);
        setExtractionMode(null);
        setActiveSheetName(null);
        setIsLoading(false);
        setError(null);
        setIsSwordUnsheathed(false);
        setTranslateTo('none');
        if (fileInputRef.current) {
            fileInputRef.current.value = '';
        }
    }, []);

    useEffect(() => {
        return () => {
            if (previewUrl) {
                URL.revokeObjectURL(previewUrl);
            }
        };
    }, [previewUrl]);

    useEffect(() => {
        const handleFocus = () => {
            setTimeout(() => {
                if (fileInputRef.current && !fileInputRef.current.files?.length) {
                    setIsSwordUnsheathed(false); 
                }
            }, 300);
        };

        window.addEventListener('focus', handleFocus);
        return () => {
            window.removeEventListener('focus', handleFocus);
        };
    }, []);
    
    const handleFileChange = useCallback((file: File | null) => {
        if (!file) {
            setIsSwordUnsheathed(false);
            return;
        }
        
        setUploadedFile(file);
        setExtractedData({});
        setExtractedText(null);
        setExtractionMode(null);
        setActiveSheetName(null);
        setError(null);
        setIsLoading(false);
        if (previewUrl) URL.revokeObjectURL(previewUrl);
        setPreviewUrl(null);

        const fileName = file.name.toLowerCase();
        const fileMimeType = file.type;

        if (fileMimeType.startsWith('image/')) {
            setFileType('image');
            const url = URL.createObjectURL(file);
            setPreviewUrl(url);
        } else if (fileMimeType.includes('sheet') || fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
            setFileType('excel');
        } else if (fileMimeType === 'application/pdf' || fileName.endsWith('.pdf')) {
            setFileType('pdf');
        } else if (
            fileMimeType === 'application/msword' ||
            fileMimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
            fileName.endsWith('.doc') ||
            fileName.endsWith('.docx')
        ) {
            setFileType('word');
        } else {
            setError('Unsupported file type. Please upload an image, Excel, Word, or PDF file.');
        }
    }, [previewUrl]);

    const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
        handleFileChange(event.target.files ? event.target.files[0] : null);
    };

    const handleDragOver = (event: React.DragEvent<HTMLDivElement>) => {
        event.preventDefault();
        setIsDragging(true);
    };

    const handleDragLeave = (event: React.DragEvent<HTMLDivElement>) => {
        event.preventDefault();
        setIsDragging(false);
    };

    const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
        event.preventDefault();
        setIsDragging(false);
        setIsSwordUnsheathed(true);
        handleFileChange(event.dataTransfer.files ? event.dataTransfer.files[0] : null);
    };

    const handleUploadAreaClick = () => {
        if(isSwordUnsheathed) return;
        setIsSwordUnsheathed(true);
        setTimeout(() => fileInputRef.current?.click(), 100);
    };

    const fileToBase64 = (file: File): Promise<string> => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = () => {
                const result = reader.result as string;
                resolve(result.split(',')[1]);
            };
            reader.onerror = (error) => reject(error);
        });
    };

    const handleProcessForExcel = async () => {
        if (!uploadedFile) {
            setError('Please upload a file first.');
            return;
        }
        if (!apiKey) {
            setError('Please set your API key first.');
            setIsModalOpen(true);
            return;
        }

        setIsLoading(true);
        setError(null);
        setExtractedData({});
        setExtractedText(null);
        setExtractionMode('excel');

        try {
            const ai = new GoogleGenAI({ apiKey: apiKey });
            const base64Data = await fileToBase64(uploadedFile);

            const filePart = {
                inlineData: {
                    mimeType: uploadedFile.type,
                    data: base64Data,
                },
            };
            
            let prompt = "Extract all text and tabular data from this document. Structure it as a JSON object with a single key 'data'. The value of 'data' should be an array of arrays, where each inner array represents a row of data. Maintain the original row and column structure of any tables found.";
            
            if (translateTo !== 'none') {
                 prompt += ` After extracting, translate all the text to ${translateTo}.`;
            }

            const textPart = { text: prompt };

            const response = await ai.models.generateContent({
                model: 'gemini-2.5-flash',
                contents: { parts: [filePart, textPart] },
                config: {
                    responseMimeType: 'application/json',
                    responseSchema: {
                        type: Type.OBJECT,
                        properties: {
                            data: {
                                type: Type.ARRAY,
                                items: {
                                    type: Type.ARRAY,
                                    items: { type: Type.STRING }
                                }
                            }
                        },
                        required: ['data']
                    }
                }
            });
            
            const jsonResponse = JSON.parse(response.text);
            const data = jsonResponse.data;
            
            if (data && Array.isArray(data)) {
                 setExtractedData({ 'Extracted Data': data });
                 setActiveSheetName('Extracted Data');
            } else {
                 throw new Error("Invalid data format received from API.");
            }

        } catch (err: any) {
            console.error(err);
            setError(`An error occurred: ${err.message || 'Please try again.'}`);
        } finally {
            setIsLoading(false);
        }
    };

    const handleProcessForWord = async () => {
        if (!uploadedFile) {
            setError('Please upload a file first.');
            return;
        }
        if (!apiKey) {
            setError('Please set your API key first.');
            setIsModalOpen(true);
            return;
        }

        setIsLoading(true);
        setError(null);
        setExtractedData({});
        setExtractedText(null);
        setExtractionMode('word');

        try {
            const ai = new GoogleGenAI({ apiKey: apiKey });
            const base64Data = await fileToBase64(uploadedFile);

            const filePart = {
                inlineData: {
                    mimeType: uploadedFile.type,
                    data: base64Data,
                },
            };

            let prompt = "Extract all content from this document. Preserve the original layout, including headings, paragraphs, lists, and tables as best as possible. The output should be a single block of formatted text. Maintain spacing and line breaks.";

            if (translateTo !== 'none') {
                prompt += ` After extracting, translate all the content to ${translateTo}.`;
            }

            const textPart = { text: prompt };

            const response = await ai.models.generateContent({
                model: 'gemini-2.5-flash',
                contents: { parts: [filePart, textPart] },
            });

            setExtractedText(response.text);

        } catch (err: any) {
            console.error(err);
            setError(`An error occurred: ${err.message || 'Please try again.'}`);
        } finally {
            setIsLoading(false);
        }
    };
    
    const processExcelFile = async () => {
         if (!uploadedFile) return;

        setIsLoading(true);
        setError(null);
        setExtractedData({});
        setExtractedText(null);
        
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const data = e.target?.result;
                const workbook = XLSX.read(data, { type: 'binary' });
                setExtractionMode('excel');

                if (translateTo === 'none') {
                    const sheetData: {[sheetName: string]: string[][]} = {};
                    workbook.SheetNames.forEach((sheetName: string) => {
                        const worksheet = workbook.Sheets[sheetName];
                        const json: string[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        sheetData[sheetName] = json;
                    });
                    setExtractedData(sheetData);
                    setActiveSheetName(workbook.SheetNames[0]);
                } else {
                    if (!apiKey) {
                        setError('Please set your API key to use the translation feature.');
                        setIsModalOpen(true);
                        setIsLoading(false);
                        return;
                    }
                    const ai = new GoogleGenAI({ apiKey: apiKey });

                    const translationPromises = workbook.SheetNames.map(async (sheetName: string): Promise<string[][] | null> => {
                        const worksheet = workbook.Sheets[sheetName];
                        const jsonData: string[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                        if (jsonData.length === 0) {
                            return null;
                        }

                        const prompt = `Translate the following tabular data into ${translateTo}. The data is a JSON array of arrays. Each inner array is a row. Maintain the exact same row and column structure. Return a JSON object with a single key 'translatedData' which contains the translated array of arrays. Data: ${JSON.stringify(jsonData)}`;

                        const response = await ai.models.generateContent({
                            model: 'gemini-2.5-flash',
                            contents: prompt,
                            config: {
                                responseMimeType: 'application/json',
                                responseSchema: {
                                    type: Type.OBJECT,
                                    properties: {
                                        translatedData: {
                                            type: Type.ARRAY,
                                            items: {
                                                type: Type.ARRAY,
                                                items: { type: Type.STRING }
                                            }
                                        }
                                    },
                                    required: ['translatedData']
                                }
                            }
                        });
                        
                        const jsonResponse = JSON.parse(response.text);
                        const translatedData = jsonResponse.translatedData;

                        if (translatedData && Array.isArray(translatedData)) {
                           return translatedData;
                        }
                        return null;
                    });
                    
                    const settledResults = await Promise.allSettled(translationPromises);
                    const translatedWorkbookData: { [sheetName: string]: string[][] } = {};
                    let sheetIndex = 1;

                    settledResults.forEach(result => {
                        if (result.status === 'fulfilled' && result.value) {
                            translatedWorkbookData[`Sheet ${sheetIndex}`] = result.value;
                            sheetIndex++;
                        } else if (result.status === 'rejected') {
                            console.error("A sheet failed to translate:", result.reason);
                        }
                    });
                    
                    setExtractedData(translatedWorkbookData);
                    if (Object.keys(translatedWorkbookData).length > 0) {
                        setActiveSheetName('Sheet 1');
                    }
                }
            } catch (err: any) {
                console.error(err);
                setError(`An error occurred: ${err.message || 'Please try again.'}`);
            } finally {
                setIsLoading(false);
            }
        };
        reader.readAsBinaryString(uploadedFile);
    };

    const handleDownload = () => {
        const wb = XLSX.utils.book_new();
        Object.keys(extractedData).forEach(sheetName => {
            const ws = XLSX.utils.aoa_to_sheet(extractedData[sheetName]);
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
        });
        const fileName = `extracted_data_${new Date().toISOString()}.xlsx`;
        XLSX.writeFile(wb, fileName);
    };

    const handleWordDownload = () => {
        let htmlString = '';
        const docTitle = "Extracted Data";

        if (extractionMode === 'word' && extractedText) {
            const escapedText = extractedText
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&#039;");

            htmlString = `
                <!DOCTYPE html><html><head><meta charset="UTF-8"><title>${docTitle}</title>
                <style>
                    body { font-family: Arial, sans-serif; font-size: 12pt; }
                    pre { white-space: pre-wrap; word-wrap: break-word; font-family: 'Courier New', Courier, monospace; }
                </style>
                </head><body><pre>${escapedText}</pre></body></html>
            `;
        } else if (extractionMode === 'excel' && Object.keys(extractedData).length > 0) {
            htmlString = `
                <!DOCTYPE html><html><head><meta charset="UTF-8"><title>${docTitle}</title>
                <style>
                    body { font-family: Arial, sans-serif; }
                    table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
                    th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; }
                    th { background-color: #f2f2f2; }
                    h2 { color: #333; }
                </style>
                </head><body>
            `;

            Object.keys(extractedData).forEach(sheetName => {
                htmlString += `<h2>${sheetName}</h2>`;
                const sheetData = extractedData[sheetName];
                if (sheetData && sheetData.length > 0) {
                    htmlString += '<table>';
                    if (sheetData[0] && sheetData[0].length > 0) {
                        htmlString += '<thead><tr>';
                        sheetData[0].forEach(header => {
                            htmlString += `<th>${header || ''}</th>`;
                        });
                        htmlString += '</tr></thead>';
                    }
                    
                    htmlString += '<tbody>';
                    sheetData.slice(1).forEach(row => {
                        htmlString += '<tr>';
                        row.forEach(cell => {
                            htmlString += `<td>${cell || ''}</td>`;
                        });
                        htmlString += '</tr>';
                    });
                    htmlString += '</tbody></table>';
                }
            });

            htmlString += '</body></html>';
        } else {
            return; // Nothing to download
        }

        const blob = new Blob([`\ufeff${htmlString}`], { type: 'application/msword' });
        
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `extracted_data_${new Date().toISOString()}.doc`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(link.href);
    };

    const isAiProcessingDisabled = !apiKey || isLoading;

    return (
        <div className="container">
            {isModalOpen && <ApiKeyModal onApiKeySubmit={handleApiKeySubmit} />}
            <div className="header">
                 <h1>Data Alchemist</h1>
                <p>Unleash the power of AI to transform documents into structured data.</p>
            </div>

            <div className="made-by-logo">
                Made by SiMon Yue De
                <p className="made-by-date">30/09/2025</p>
            </div>
            
            {!uploadedFile && (
                 <div
                    className={`sword-upload-container ${isDragging ? 'drag-over' : ''} ${!apiKey ? 'disabled' : ''}`}
                    onClick={apiKey ? handleUploadAreaClick : () => setIsModalOpen(true)}
                    onDragOver={apiKey ? handleDragOver : (e) => e.preventDefault()}
                    onDragLeave={apiKey ? handleDragLeave : (e) => e.preventDefault()}
                    onDrop={apiKey ? handleDrop : (e) => e.preventDefault()}
                    role="button"
                    tabIndex={apiKey ? 0 : -1}
                    aria-label="Upload file by clicking the sword"
                    aria-disabled={!apiKey}
                >
                    <div className={`sword ${isSwordUnsheathed ? 'unsheathed' : ''}`}>
                         <svg className="sword-hilt" viewBox="0 0 90 50" preserveAspectRatio="none">
                             <defs>
                                 <linearGradient id="hiltGrad3" x1="0" y1="0" x2="0" y2="1">
                                     <stop offset="0%" stopColor="#FCEABB"/>
                                     <stop offset="50%" stopColor="#F8C688"/>
                                     <stop offset="100%" stopColor="#B8860B"/>
                                 </linearGradient>
                             </defs>
                              {/* Guard */}
                             <path d="M40,5 C60,5,75,15,80,20 H85 V30 H80 C75,35,60,45,40,45 L30,25 L40,5z" fill="url(#hiltGrad3)"/>
                             <circle cx="60" cy="25" r="4" fill="#B22222"/>
                             {/* Hilt */}
                             <path d="M5,18 H45 V32 H5 C-5,32,-5,18,5,18z" fill="#654321"/>
                             <rect x="10" y="19" width="30" height="12" fill="#8B4513"/>
                             {/* Pommel */}
                             <circle cx="5" cy="25" r="8" fill="url(#hiltGrad3)"/>
                         </svg>
                         <svg className="sword-blade" viewBox="0 0 410 20" preserveAspectRatio="none">
                             <defs>
                                 <linearGradient id="bladeGrad2" x1="0" y1="0" x2="0" y2="1">
                                     <stop offset="0%" stopColor="#FFFFFF"/>
                                     <stop offset="50%" stopColor="#D3D3D3"/>
                                     <stop offset="100%" stopColor="#E5E5E5"/>
                                 </linearGradient>
                                 <filter id="bladeShine2">
                                     <feGaussianBlur stdDeviation="1"/>
                                 </filter>
                             </defs>
                             <path d="M0,10 L15,0 H400 C402.8,0,405,2.2,405,5v10 c0,2.8-2.2,5-5,5 H15 L0,10z" fill="url(#bladeGrad2)" filter="url(#bladeShine2)"/>
                             <path d="M15,2 H400 C401.7,2,403,3.3,403,5v10 c0,1.7-1.3,3-3,3 H15 L5,10 L15,2z" fill="#FFFFFF" opacity="0.8"/>
                         </svg>
                    </div>
 
                     <svg className="sword-scabbard" viewBox="0 0 425 32" preserveAspectRatio="none">
                         <defs>
                             <linearGradient id="scabbardGrad2" x1="0" y1="0" x2="0" y2="1">
                                 <stop offset="0%" stopColor="#4A2A1E"/>
                                 <stop offset="50%" stopColor="#3A1F16"/>
                                 <stop offset="100%" stopColor="#4A2A1E"/>
                             </linearGradient>
                              <linearGradient id="hiltGrad2" x1="0" y1="0" x2="0" y2="1">
                                 <stop offset="0%" stopColor="#FCEABB"/>
                                 <stop offset="50%" stopColor="#F8C688"/>
                                 <stop offset="100%" stopColor="#B8860B"/>
                             </linearGradient>
                         </defs>
                         <path d="M415,1 H10 C4.5,1,0,5.5,0,11v10 C0,26.5,4.5,31,10,31 H415 C420.5,31,425,26.5,425,21 V11 C425,5.5,420.5,1,415,1z" fill="url(#scabbardGrad2)"/>
                         <path d="M410,3 H422 V29 H410z" fill="url(#hiltGrad2)"/>
                         <path d="M5,8 H20 V24 H5z" fill="url(#hiltGrad2)"/>
                         <path d="M412,8 a 2,2 0 0,1 4,0 v16 a 2,2 0 0,1 -4,0 z" fill="#8B4513" opacity="0.5"/>
                         <path d="M7,13 a 1,1 0 0,1 2,0 v6 a 1,1 0 0,1 -2,0 z" fill="#8B4513" opacity="0.5"/>
                     </svg>

                    <p className="sword-prompt">{apiKey ? 'Click the sword to upload a file' : 'Please provide an API key to begin'}</p>
                </div>
            )}

             <input
                id="file-input"
                type="file"
                ref={fileInputRef}
                onChange={handleFileUpload}
                accept="image/*,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,.doc,.docx,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/pdf"
                disabled={!apiKey}
            />
            
            {uploadedFile && (
                <div className="file-info">
                    {previewUrl && fileType === 'image' && (
                        <div className="image-preview">
                            <img src={previewUrl} alt="Uploaded preview" />
                        </div>
                    )}
                    <div className="actions-container">
                        <div className="translate-option">
                            <label htmlFor="translate-select">Tùy chọn dịch thuật:</label>
                            <select id="translate-select" className="translate-select" value={translateTo} onChange={(e) => setTranslateTo(e.target.value)} disabled={!apiKey}>
                                <option value="none">Giữ ngôn ngữ gốc</option>
                                <option value="English">Tiếng Anh</option>
                                <option value="Vietnamese">Tiếng Việt</option>
                                <option value="Chinese">Tiếng Trung</option>
                            </select>
                        </div>
                        <div className="button-group">
                            {(fileType === 'image' || fileType === 'pdf' || fileType === 'word') && (
                                <>
                                    <button onClick={handleProcessForExcel} className="btn btn-primary" disabled={isAiProcessingDisabled}>
                                        {isLoading && extractionMode === 'excel' ? "Đang xử lý..." : `Trích xuất ${translateTo !== 'none' ? '& Dịch ' : ''}cho Excel`}
                                    </button>
                                    <button onClick={handleProcessForWord} className="btn btn-primary" disabled={isAiProcessingDisabled}>
                                        {isLoading && extractionMode === 'word' ? "Đang xử lý..." : `Trích xuất ${translateTo !== 'none' ? '& Dịch ' : ''}cho Word`}
                                    </button>
                                </>
                            )}
                            {fileType === 'excel' && (
                                <button onClick={processExcelFile} className="btn btn-primary" disabled={isLoading}>
                                    {isLoading ? "Đang xử lý..." : (translateTo === 'none' ? 'Xem dữ liệu Excel' : 'Dịch & Xem dữ liệu')}
                                </button>
                            )}

                            {Object.keys(extractedData).length > 0 && (
                                <>
                                    <button onClick={handleDownload} className="btn btn-secondary">Tải xuống file Excel</button>
                                    <button onClick={handleWordDownload} className="btn btn-secondary">Tải xuống file Word</button>
                                </>
                            )}
                            {extractedText && (
                                <button onClick={handleWordDownload} className="btn btn-secondary">Tải xuống file Word</button>
                            )}
                            <button onClick={resetState} className="btn btn-tertiary">Tải lên tệp khác</button>
                        </div>
                    </div>
                </div>
            )}

            {isLoading && (
                <div className="loading-container">
                    <div className="cauldron-loader">
                        <div className="cauldron-shadow"></div>
                        <div className="cauldron">
                            <div className="cauldron-liquid"></div>
                            <div className="cauldron-bubble b1"></div>
                            <div className="cauldron-bubble b2"></div>
                            <div className="cauldron-bubble b3"></div>
                            <div className="cauldron-bubble b4"></div>
                            <div className="cauldron-bubble b5"></div>
                            <div className="cauldron-steam s1"></div>
                            <div className="cauldron-steam s2"></div>
                            <div className="cauldron-steam s3"></div>
                        </div>
                    </div>
                    <p className="loading-text">Brewing your data...</p>
                </div>
            )}
            {error && <div className="error-message">{error}</div>}
            
            {extractionMode === 'word' && extractedText && (
                <div className="output-section">
                    <div className="text-results-container">
                        <pre>{extractedText}</pre>
                    </div>
                </div>
            )}

            {extractionMode === 'excel' && Object.keys(extractedData).length > 0 && (
                <div className="output-section">
                     <div className="sheet-tabs-container">
                        {Object.keys(extractedData).map(sheetName => (
                            <button
                                key={sheetName}
                                className={`sheet-tab ${sheetName === activeSheetName ? 'active' : ''}`}
                                onClick={() => setActiveSheetName(sheetName)}
                            >
                                {sheetName}
                            </button>
                        ))}
                    </div>

                    {activeSheetName && extractedData[activeSheetName] && (
                        <div className="results-container">
                            <table className="results-table">
                                <thead>
                                    <tr>
                                        {extractedData[activeSheetName][0]?.map((header, index) => (
                                            <th key={index}>{header}</th>
                                        ))}
                                    </tr>
                                </thead>
                                <tbody>
                                    {extractedData[activeSheetName].slice(1).map((row, rowIndex) => (
                                        <tr key={rowIndex}>
                                            {row.map((cell, cellIndex) => (
                                                <td key={cellIndex}>{cell}</td>
                                            ))}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                </div>
            )}
        </div>
    );
};

const root = createRoot(document.getElementById('root')!);
root.render(<App />);