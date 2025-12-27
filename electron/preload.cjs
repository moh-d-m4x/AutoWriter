const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    // Template operations
    getTemplatePath: () => ipcRenderer.invoke('get-template-path'),
    readTemplate: () => ipcRenderer.invoke('read-template'),

    // Document operations
    saveDocument: (options) => ipcRenderer.invoke('save-document', options),
    saveMultipleImages: (images, defaultBaseName) => ipcRenderer.invoke('save-multiple-images', { images, defaultBaseName }),

    // Image operations
    selectImage: () => ipcRenderer.invoke('select-image'),
    selectImages: () => ipcRenderer.invoke('select-images'),
    selectFiles: () => ipcRenderer.invoke('select-files'), // All supported file types
    readImage: (path) => ipcRenderer.invoke('read-image', path),

    // Conversion operations (for PDF/Image export matching Word template)
    convertToPdf: (formData) => ipcRenderer.invoke('convert-to-pdf', { formData }),
    convertToImage: (formData) => ipcRenderer.invoke('convert-to-image', { formData }),

    // Word-based conversion (1:1 matching with template)
    convertDocxToPdfWord: (docxBuffer) => ipcRenderer.invoke('convert-docx-to-pdf-word', { docxBuffer }),
    convertDocxToImageWord: (docxBuffer, combinePages, subjectName) => ipcRenderer.invoke('convert-docx-to-image-word', { docxBuffer, combinePages, subjectName }),

    // Direct PDF to Image conversion (for dropped/shared PDF files)
    convertPdfToImage: (pdfBuffer, combinePages) => ipcRenderer.invoke('convert-pdf-to-image', { pdfBuffer, combinePages }),

    // Cancel export operation
    cancelExport: () => ipcRenderer.invoke('cancel-export'),

    // Open URL in system's default browser
    openExternal: (url) => ipcRenderer.invoke('open-external', url),

    // Listen for files opened via "Open with"
    onFileOpened: (callback) => {
        ipcRenderer.on('file-opened', (event, fileData) => callback(fileData));
        // Return cleanup function
        return () => ipcRenderer.removeAllListeners('file-opened');
    }
});
