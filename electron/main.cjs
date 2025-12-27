const { app, BrowserWindow, ipcMain, dialog, Menu, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { spawn } = require('child_process');

// Helper function to get the correct path for unpacked asar files
// When app is packaged, files unpacked via asarUnpack go to app.asar.unpacked folder
function getUnpackedPath(relativePath) {
    if (app.isPackaged) {
        // In packaged app, use app.asar.unpacked path
        return path.join(__dirname.replace('app.asar', 'app.asar.unpacked'), relativePath);
    } else {
        // In development, use normal path
        return path.join(__dirname, relativePath);
    }
}

let mainWindow;
let currentExportProcess = null; // Track current export process for cancellation
let exportCancelled = false; // Track if export was cancelled
let pendingFilePath = null; // File path from "Open with" command

// Supported file extensions for "Open with"
const SUPPORTED_EXTENSIONS = ['.pdf', '.docx', '.doc', '.pptx', '.ppt', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'];

// Check if a file path is supported
function isSupportedFile(filePath) {
    if (!filePath) return false;
    const ext = path.extname(filePath).toLowerCase();
    return SUPPORTED_EXTENSIONS.includes(ext);
}

// Get file path from command line arguments
function getFileFromArgs(args) {
    // Skip the first arg (electron executable) and look for file paths
    for (let i = 1; i < args.length; i++) {
        const arg = args[i];
        // Skip flags and special arguments
        if (arg.startsWith('-') || arg.startsWith('--')) continue;
        // Check if it's a valid file path
        if (fs.existsSync(arg) && isSupportedFile(arg)) {
            return arg;
        }
    }
    return null;
}

// Request single instance lock - prevents multiple app instances
const gotTheLock = app.requestSingleInstanceLock();

if (!gotTheLock) {
    // Another instance is running, quit this one
    app.quit();
} else {
    // Handle second-instance event (when user opens another file while app is running)
    app.on('second-instance', (event, commandLine, workingDirectory) => {
        // Focus the main window
        if (mainWindow) {
            if (mainWindow.isMinimized()) mainWindow.restore();
            mainWindow.focus();

            // Check for file in command line
            const filePath = getFileFromArgs(commandLine);
            if (filePath && mainWindow.webContents) {
                // Send file to renderer
                sendFileToRenderer(filePath);
            }
        }
    });
}

// Send file to renderer for processing
function sendFileToRenderer(filePath) {
    if (!mainWindow || !mainWindow.webContents) return;

    try {
        const buffer = fs.readFileSync(filePath);
        const base64 = buffer.toString('base64');
        const name = path.basename(filePath);

        mainWindow.webContents.send('file-opened', {
            name: name,
            data: base64
        });
    } catch (err) {
        console.error('Failed to read file:', err);
    }
}

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1200,
        height: 900,
        minWidth: 900,
        minHeight: 700,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, 'preload.cjs')
        },
        icon: path.join(__dirname, 'icon.png'),
        title: 'كاتب المستندات'
    });

    // Load the app
    if (process.env.NODE_ENV === 'development' || !app.isPackaged) {
        mainWindow.loadURL('http://localhost:5173');
        // DevTools: press Ctrl+Shift+I to open
    } else {
        mainWindow.loadFile(path.join(__dirname, '../dist/index.html'));
    }

    // When the window is ready, send any pending file
    mainWindow.webContents.on('did-finish-load', () => {
        if (pendingFilePath) {
            // Small delay to ensure React app is ready
            setTimeout(() => {
                sendFileToRenderer(pendingFilePath);
                pendingFilePath = null;
            }, 500);
        }
    });
}

app.whenReady().then(() => {
    // Check for file in startup arguments
    pendingFilePath = getFileFromArgs(process.argv);

    // Remove the menu bar
    Menu.setApplicationMenu(null);
    createWindow();
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
    }
});

// IPC Handlers

// Get template path
ipcMain.handle('get-template-path', () => {
    return getUnpackedPath('../ref/template.docx');
});

// Read template file
ipcMain.handle('read-template', async () => {
    const templatePath = getUnpackedPath('../ref/template.docx');
    return fs.readFileSync(templatePath);
});

// Save generated document
ipcMain.handle('save-document', async (event, { buffer, defaultName, filters }) => {
    const result = await dialog.showSaveDialog(mainWindow, {
        defaultPath: defaultName,
        filters: filters
    });

    if (!result.canceled && result.filePath) {
        fs.writeFileSync(result.filePath, Buffer.from(buffer));
        return { success: true, path: result.filePath };
    }
    return { success: false };
});

// Save multiple images with auto-numbering
ipcMain.handle('save-multiple-images', async (event, { images, defaultBaseName }) => {
    // Sanitize filename - replace Windows unsupported characters with dash
    const sanitizedBaseName = defaultBaseName.replace(/[\\/:*?"<>|]/g, '-');

    // Show save dialog once - user picks folder and base name
    const result = await dialog.showSaveDialog(mainWindow, {
        defaultPath: `${sanitizedBaseName}.png`,
        filters: [{ name: 'PNG Image', extensions: ['png'] }]
    });

    if (!result.canceled && result.filePath) {
        // Extract directory and base name from the chosen path
        const chosenDir = path.dirname(result.filePath);
        const chosenFileName = path.basename(result.filePath, '.png');
        // Remove trailing _1 if present to get clean base name
        const baseName = chosenFileName.replace(/_\d+$/, '');

        // Save all images with page numbers
        for (let i = 0; i < images.length; i++) {
            const pageNum = i + 1;
            const fileName = `${baseName}_${pageNum}.png`;
            const filePath = path.join(chosenDir, fileName);
            fs.writeFileSync(filePath, Buffer.from(images[i].buffer));
        }

        return { success: true, count: images.length, directory: chosenDir };
    }
    return { success: false };
});

// Cancel current export operation
ipcMain.handle('cancel-export', async () => {
    exportCancelled = true;
    if (currentExportProcess) {
        try {
            // Kill the PowerShell process and its children
            spawn('taskkill', ['/pid', currentExportProcess.pid.toString(), '/f', '/t']);
            currentExportProcess = null;
        } catch (err) {
            console.error('Error killing export process:', err);
        }
    }
    return { success: true };
});

// Convert DOCX to PDF using MS Word (Windows only)
ipcMain.handle('convert-docx-to-pdf-word', async (event, { docxBuffer }) => {
    // Reset cancelled flag at start
    exportCancelled = false;

    return new Promise(async (resolve, reject) => {
        try {
            const tempDir = os.tmpdir();
            const timestamp = Date.now();
            const tempDocxPath = path.join(tempDir, `autowriter_temp_${timestamp}.docx`);
            const tempPdfPath = path.join(tempDir, `autowriter_temp_${timestamp}.pdf`);

            // Write DOCX to temp file
            fs.writeFileSync(tempDocxPath, Buffer.from(docxBuffer));

            // Path to PowerShell script
            const scriptPath = getUnpackedPath('word-converter.ps1');

            // Run PowerShell script
            const ps = spawn('powershell.exe', [
                '-ExecutionPolicy', 'Bypass',
                '-File', scriptPath,
                '-InputPath', tempDocxPath,
                '-OutputPath', tempPdfPath,
                '-Format', 'pdf'
            ]);

            // Track current process for cancellation
            currentExportProcess = ps;

            let stdout = '';
            let stderr = '';

            ps.stdout.on('data', (data) => {
                stdout += data.toString();
            });

            ps.stderr.on('data', (data) => {
                stderr += data.toString();
            });

            ps.on('close', (code) => {
                // Clear current process reference
                currentExportProcess = null;

                // Cleanup temp docx
                try { fs.unlinkSync(tempDocxPath); } catch (e) { }

                // Check if cancelled
                if (exportCancelled) {
                    try { fs.unlinkSync(tempPdfPath); } catch (e) { }
                    reject(new Error('Export cancelled'));
                    return;
                }

                if (code === 0 && fs.existsSync(tempPdfPath)) {
                    const pdfBuffer = fs.readFileSync(tempPdfPath);
                    try { fs.unlinkSync(tempPdfPath); } catch (e) { }
                    resolve({ success: true, buffer: Array.from(pdfBuffer) });
                } else {
                    try { fs.unlinkSync(tempPdfPath); } catch (e) { }
                    reject(new Error(`Word conversion failed: ${stderr || stdout}`));
                }
            });

            ps.on('error', (err) => {
                try { fs.unlinkSync(tempDocxPath); } catch (e) { }
                reject(new Error(`Failed to start PowerShell: ${err.message}`));
            });

            // Timeout after 60 seconds
            setTimeout(() => {
                ps.kill();
                try { fs.unlinkSync(tempDocxPath); } catch (e) { }
                try { fs.unlinkSync(tempPdfPath); } catch (e) { }
                reject(new Error('Word conversion timeout'));
            }, 60000);

        } catch (err) {
            reject(err);
        }
    });
});

// Convert DOCX to Image using MS Word + pdf-poppler (Windows only)
// Also handles direct PDF input (detected via subjectName extension)
// Also handles PowerPoint files (.pptx, .ppt)
ipcMain.handle('convert-docx-to-image-word', async (event, { docxBuffer, combinePages, subjectName }) => {
    // Reset cancelled flag at start
    exportCancelled = false;

    // Detect file type from extension
    const lowerName = subjectName ? subjectName.toLowerCase() : '';
    const isPdf = lowerName.endsWith('.pdf');
    const isPowerPoint = lowerName.endsWith('.pptx') || lowerName.endsWith('.ppt');

    return new Promise(async (resolve, reject) => {
        try {
            const tempDir = os.tmpdir();
            const timestamp = Date.now();

            if (isPdf) {
                // Direct PDF handling - skip Word conversion
                const tempPdfPath = path.join(tempDir, `autowriter_temp_${timestamp}.pdf`);

                // Write PDF to temp file
                fs.writeFileSync(tempPdfPath, Buffer.from(docxBuffer));

                // Convert PDF to images directly
                try {
                    const pdfPoppler = require('pdf-poppler');

                    const outputDir = path.join(tempDir, `autowriter_images_${timestamp}`);
                    fs.mkdirSync(outputDir, { recursive: true });

                    const opts = {
                        format: 'png',
                        out_dir: outputDir,
                        out_prefix: 'page',
                        scale: 2048
                    };

                    await pdfPoppler.convert(tempPdfPath, opts);

                    const files = fs.readdirSync(outputDir)
                        .filter(f => f.endsWith('.png'))
                        .sort();

                    if (files.length === 0) {
                        throw new Error('No image files generated');
                    }

                    if (combinePages) {
                        if (files.length === 1) {
                            const imagePath = path.join(outputDir, files[0]);
                            const pngBuffer = fs.readFileSync(imagePath);

                            try { fs.unlinkSync(imagePath); } catch (e) { }
                            try { fs.rmdirSync(outputDir); } catch (e) { }
                            try { fs.unlinkSync(tempPdfPath); } catch (e) { }

                            resolve({ success: true, buffer: Array.from(pngBuffer) });
                        } else {
                            const { createCanvas, loadImage } = require('canvas');

                            const images = [];
                            let totalHeight = 0;
                            let maxWidth = 0;

                            for (const file of files) {
                                const imagePath = path.join(outputDir, file);
                                const img = await loadImage(imagePath);
                                images.push(img);
                                totalHeight += img.height;
                                maxWidth = Math.max(maxWidth, img.width);
                            }

                            const canvas = createCanvas(maxWidth, totalHeight);
                            const ctx = canvas.getContext('2d');
                            ctx.fillStyle = 'white';
                            ctx.fillRect(0, 0, maxWidth, totalHeight);

                            let yOffset = 0;
                            for (const img of images) {
                                ctx.drawImage(img, 0, yOffset);
                                yOffset += img.height;
                            }

                            const pngBuffer = canvas.toBuffer('image/png');

                            for (const file of files) {
                                try { fs.unlinkSync(path.join(outputDir, file)); } catch (e) { }
                            }
                            try { fs.rmdirSync(outputDir); } catch (e) { }
                            try { fs.unlinkSync(tempPdfPath); } catch (e) { }

                            resolve({ success: true, buffer: Array.from(pngBuffer) });
                        }
                    } else {
                        const imageBuffers = [];

                        for (let i = 0; i < files.length; i++) {
                            const imagePath = path.join(outputDir, files[i]);
                            const pngBuffer = fs.readFileSync(imagePath);
                            imageBuffers.push({
                                buffer: Array.from(pngBuffer),
                                pageNum: i + 1
                            });
                        }

                        for (const file of files) {
                            try { fs.unlinkSync(path.join(outputDir, file)); } catch (e) { }
                        }
                        try { fs.rmdirSync(outputDir); } catch (e) { }
                        try { fs.unlinkSync(tempPdfPath); } catch (e) { }

                        resolve({ success: true, images: imageBuffers });
                    }
                } catch (pdfErr) {
                    console.error('PDF to image conversion error:', pdfErr);
                    try { fs.unlinkSync(tempPdfPath); } catch (e) { }
                    resolve({ success: false, error: `PDF conversion failed: ${pdfErr.message}` });
                }
                return;
            }

            // DOCX/PPTX handling - use Word/PowerPoint to convert to PDF first
            // Determine the correct file extension
            let inputExt = '.docx';
            if (isPowerPoint) {
                inputExt = lowerName.endsWith('.ppt') ? '.ppt' : '.pptx';
            } else if (lowerName.endsWith('.doc')) {
                inputExt = '.doc';
            }

            const tempInputPath = path.join(tempDir, `autowriter_temp_${timestamp}${inputExt}`);
            const tempPdfPath = path.join(tempDir, `autowriter_temp_${timestamp}.pdf`);

            // Write input file to temp
            fs.writeFileSync(tempInputPath, Buffer.from(docxBuffer));

            // Path to PowerShell script
            const scriptPath = getUnpackedPath('word-converter.ps1');

            // First convert to PDF using Word
            const ps = spawn('powershell.exe', [
                '-ExecutionPolicy', 'Bypass',
                '-File', scriptPath,
                '-InputPath', tempInputPath,
                '-OutputPath', tempPdfPath,
                '-Format', 'pdf'
            ]);

            // Track current process for cancellation
            currentExportProcess = ps;

            let stdout = '';
            let stderr = '';

            ps.stdout.on('data', (data) => {
                stdout += data.toString();
            });

            ps.stderr.on('data', (data) => {
                stderr += data.toString();
            });

            ps.on('close', async (code) => {
                // Clear current process reference
                currentExportProcess = null;

                // Cleanup temp input file
                try { fs.unlinkSync(tempInputPath); } catch (e) { }

                // Check if cancelled
                if (exportCancelled) {
                    try { fs.unlinkSync(tempPdfPath); } catch (e) { }
                    reject(new Error('Export cancelled'));
                    return;
                }

                if (code === 0 && fs.existsSync(tempPdfPath)) {
                    // Convert PDF to image using pdf-poppler
                    try {
                        const pdfPoppler = require('pdf-poppler');

                        // Output directory for images
                        const outputDir = path.join(tempDir, `autowriter_images_${timestamp}`);
                        fs.mkdirSync(outputDir, { recursive: true });

                        // Convert options - export ALL pages
                        const opts = {
                            format: 'png',
                            out_dir: outputDir,
                            out_prefix: 'page',
                            scale: 2048 // High resolution
                            // No 'page' option = convert all pages
                        };

                        await pdfPoppler.convert(tempPdfPath, opts);

                        // Read all generated images
                        const files = fs.readdirSync(outputDir)
                            .filter(f => f.endsWith('.png'))
                            .sort(); // Sort to maintain page order

                        if (files.length === 0) {
                            throw new Error('No image files generated');
                        }

                        if (combinePages) {
                            if (files.length === 1) {
                                // Single page - just return it as buffer
                                const imagePath = path.join(outputDir, files[0]);
                                const pngBuffer = fs.readFileSync(imagePath);

                                // Cleanup
                                try { fs.unlinkSync(imagePath); } catch (e) { }
                                try { fs.rmdirSync(outputDir); } catch (e) { }
                                try { fs.unlinkSync(tempPdfPath); } catch (e) { }

                                resolve({ success: true, buffer: Array.from(pngBuffer) });
                            } else {
                                // Combine all pages into one tall image
                                const { createCanvas, loadImage } = require('canvas');

                                // Load all images and get dimensions
                                const images = [];
                                let totalHeight = 0;
                                let maxWidth = 0;

                                for (const file of files) {
                                    const imagePath = path.join(outputDir, file);
                                    const img = await loadImage(imagePath);
                                    images.push(img);
                                    totalHeight += img.height;
                                    maxWidth = Math.max(maxWidth, img.width);
                                }

                                // Create canvas for combined image
                                const canvas = createCanvas(maxWidth, totalHeight);
                                const ctx = canvas.getContext('2d');

                                // Fill with white background
                                ctx.fillStyle = 'white';
                                ctx.fillRect(0, 0, maxWidth, totalHeight);

                                // Draw each image
                                let yOffset = 0;
                                for (const img of images) {
                                    ctx.drawImage(img, 0, yOffset);
                                    yOffset += img.height;
                                }

                                // Convert to PNG buffer
                                const pngBuffer = canvas.toBuffer('image/png');

                                // Cleanup all files
                                for (const file of files) {
                                    try { fs.unlinkSync(path.join(outputDir, file)); } catch (e) { }
                                }
                                try { fs.rmdirSync(outputDir); } catch (e) { }
                                try { fs.unlinkSync(tempPdfPath); } catch (e) { }

                                resolve({ success: true, buffer: Array.from(pngBuffer) });
                            }
                        } else {
                            // Return separate images (default behavior)
                            const imageBuffers = [];

                            for (let i = 0; i < files.length; i++) {
                                const imagePath = path.join(outputDir, files[i]);
                                const pngBuffer = fs.readFileSync(imagePath);
                                imageBuffers.push({
                                    buffer: Array.from(pngBuffer),
                                    pageNum: i + 1
                                });
                            }

                            // Cleanup all files
                            for (const file of files) {
                                try { fs.unlinkSync(path.join(outputDir, file)); } catch (e) { }
                            }
                            try { fs.rmdirSync(outputDir); } catch (e) { }
                            try { fs.unlinkSync(tempPdfPath); } catch (e) { }

                            resolve({ success: true, images: imageBuffers });
                        }

                    } catch (imgErr) {
                        console.error('PDF to image conversion error:', imgErr);
                        try { fs.unlinkSync(tempPdfPath); } catch (e) { }
                        reject(new Error(`Image conversion failed: ${imgErr.message}`));
                    }
                } else {
                    try { fs.unlinkSync(tempPdfPath); } catch (e) { }
                    reject(new Error(`Word conversion failed: ${stderr || stdout}`));
                }
            });

            ps.on('error', (err) => {
                try { fs.unlinkSync(tempDocxPath); } catch (e) { }
                reject(new Error(`Failed to start PowerShell: ${err.message}`));
            });

            // Timeout after 90 seconds (longer for image)
            setTimeout(() => {
                ps.kill();
                try { fs.unlinkSync(tempDocxPath); } catch (e) { }
                try { fs.unlinkSync(tempPdfPath); } catch (e) { }
                reject(new Error('Word conversion timeout'));
            }, 90000);

        } catch (err) {
            reject(err);
        }
    });
});

// Read image file and convert to base64
ipcMain.handle('read-image', async (event, imagePath) => {
    try {
        const buffer = fs.readFileSync(imagePath);
        return buffer.toString('base64');
    } catch (error) {
        return null;
    }
});

// Select image file
ipcMain.handle('select-image', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openFile'],
        filters: [
            { name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp'] }
        ]
    });

    if (!result.canceled && result.filePaths.length > 0) {
        const imagePath = result.filePaths[0];
        const buffer = fs.readFileSync(imagePath);
        return {
            path: imagePath,
            base64: buffer.toString('base64'),
            name: path.basename(imagePath)
        };
    }
    return null;
});

// Select multiple image files
ipcMain.handle('select-images', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openFile', 'multiSelections'],
        filters: [
            { name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'] }
        ]
    });

    if (!result.canceled && result.filePaths.length > 0) {
        const images = [];
        for (const imagePath of result.filePaths) {
            const buffer = fs.readFileSync(imagePath);
            images.push({
                path: imagePath,
                base64: buffer.toString('base64'),
                name: path.basename(imagePath)
            });
        }
        return images;
    }
    return [];
});

// Select multiple files (all supported types: images, Word, PDF, PowerPoint)
ipcMain.handle('select-files', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openFile', 'multiSelections'],
        filters: [
            { name: 'All Supported Files', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'pdf', 'docx', 'doc', 'pptx', 'ppt'] },
            { name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'] },
            { name: 'PDF Documents', extensions: ['pdf'] },
            { name: 'Word Documents', extensions: ['docx', 'doc'] },
            { name: 'PowerPoint', extensions: ['pptx', 'ppt'] }
        ]
    });

    if (!result.canceled && result.filePaths.length > 0) {
        const files = [];
        for (const filePath of result.filePaths) {
            const buffer = fs.readFileSync(filePath);
            files.push({
                path: filePath,
                base64: buffer.toString('base64'),
                name: path.basename(filePath)
            });
        }
        return files;
    }
    return [];
});

// Open URL in system's default browser
ipcMain.handle('open-external', async (event, url) => {
    try {
        await shell.openExternal(url);
        return { success: true };
    } catch (error) {
        console.error('Failed to open external URL:', error);
        return { success: false, error: error.message };
    }
});

// Helper function to generate HTML template matching Word layout
function generateDocumentHtml(formData) {
    const today = new Date();
    const arabicDays = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
    const dayName = arabicDays[today.getDay()];
    const dateStr = `${today.getDate()}/${today.getMonth() + 1}/${today.getFullYear()}م`;

    // Get logo base64
    const logoSrc = formData.logoBase64 ? `data:image/png;base64,${formData.logoBase64}` : '';

    // Generate copy_to list items
    const copyToItems = formData.copy_to
        ? formData.copy_to.split('\n').filter(line => line.trim()).map(line => `<li>${line.trim()}</li>`).join('')
        : '';

    return `<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
    <meta charset="UTF-8">
    <style>
        @page { size: A4; margin: 0; }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Arial', sans-serif;
            direction: rtl;
            background: white;
            width: 210mm;
            height: 297mm;
            position: relative;
            /* Margins from Word HTML: ~42pt (1.5cm) Top/Sides, 1cm Bottom */
            padding: 1.5cm 1.5cm 1cm 1.5cm; 
            margin: 0 auto;
        }
        
        /* Page Border from Word HTML: 1.5pt solid windowtext */
        .page-border {
            position: absolute;
            top: 15pt;
            left: 15pt;
            right: 15pt;
            bottom: 15pt;
            border: 1.5pt solid #000;
            pointer-events: none;
            z-index: 10;
        }

        /* Header Layout */
        .header { 
            display: flex; 
            justify-content: space-between; 
            align-items: flex-start; 
            margin-bottom: 25px; 
            position: relative; 
            height: 120px;
        }
        
        /* Header Horizontal Line (Double Line) */
        .header::after {
            content: '';
            position: absolute;
            bottom: -5px;
            left: 0;
            right: 0;
            height: 4px;
            border-top: 1pt solid #bf9000;
            border-bottom: 2.5pt solid #bf9000;
        }

        /* Right Section (Title) */
        .header-right { 
            text-align: right; 
            width: 35%;
            color: #808080; 
            font-weight: bold; 
            padding-top: 5px;
        }
        /* Word HTML used 20pt for headers? Maybe slightly smaller for this top section */
        .header-right h1 { font-family: 'Arial', sans-serif; font-size: 14pt; margin: 0 0 6px 0; color: #666; font-weight: bold; letter-spacing: 0.5px; }
        
        /* Center Section (Logo) */
        .header-center { 
            text-align: center; 
            width: 30%; 
            display: flex;
            justify-content: center;
            align-items: center;
            padding-top: 5px;
        }
        .header-center img { width: 95px; height: 95px; object-fit: contain; }
        
        /* Left Section (Date Box) */
        .header-left { 
            width: 35%; 
            display: flex;
            justify-content: flex-end;
            padding-top: 10px;
        }
        
        .date-box {
            border: 1px solid #7f7f7f;
            padding: 4px 8px;
            width: 220px;
            font-size: 11pt;
            font-weight: bold;
            font-family: 'Arial', sans-serif;
            background-color: transparent;
        }
        
        .header-row { 
            display: flex; 
            align-items: center;
            margin-bottom: 4px;
            white-space: nowrap;
        }
        
        .header-label { 
            width: 60px; 
            display: inline-block; 
            text-align: right;
            margin-left: 8px;
            color: #333;
        }
        
        /* Recipient Section: "The Brother / ..." */
        .recipient-section { 
            display: flex; 
            flex-direction: column; 
            align-items: flex-start; 
            margin: 40px 0 30px 0; 
            /* Word HTML: 20pt */
            font-size: 20pt; 
            font-weight: bold; 
            color: #000;
        }
        .recipient-name { margin-bottom: 10px; }

        /* Greetings: "Peace be upon you..." */
        .greetings { 
            text-align: center; 
            font-size: 20pt; 
            margin: 25px 0; 
            font-weight: normal; /* Word used span styling, seemed normal weight or bold? Usually bold */
            font-weight: bold; 
        }
        
        /* Subject Name: "Motorcycle Order" */
        .subject { 
            text-align: center; 
            font-size: 20pt; 
            font-weight: bold; 
            text-decoration: underline; 
            margin: 20px 0; 
        }
        
        /* Body Content */
        .body-content { 
            font-size: 18pt; 
            line-height: 1.5; 
            text-align: justify; 
            margin: 20px 5mm; 
            min-height: 100px; 
            font-weight: normal;
        }
        
        /* Ending: "Thanks" */
        .ending { text-align: center; font-size: 18pt; margin: 40px 0; font-weight: bold; }
        
        /* Signature */
        .signature { text-align: left; font-size: 16pt; font-weight: bold; color: #c00000; margin: 50px 0 0 30mm; }
        
        /* Copy To */
        /* Recalibrate position based on new margins/font sizes */
        .copy-to { 
            position: absolute; 
            bottom: 25mm; /* ~1 inch from bottom */
            right: 25mm;  /* ~1 inch from right */
            font-size: 14pt; 
            font-weight: bold;
        }
        .copy-to h4 { margin-bottom: 5px; font-weight: bold; }
        .copy-to ul { list-style: none; padding: 0; margin: 0; }
        .copy-to li { margin-bottom: 3px; }
        
        .watermark { 
            position: absolute; 
            top: 50%; 
            left: 50%; 
            transform: translate(-50%, -50%); 
            opacity: 0.08; 
            pointer-events: none; 
            z-index: 0; 
        }
        .watermark img { width: 400px; height: 400px; }
    </style>
</head>
<body>
    <div class="page-border"></div>
    ${logoSrc ? `<div class="watermark"><img src="${logoSrc}" alt=""></div>` : ''}
    
    <div class="header">
        <!-- Title Section (Right in RTL) -->
        <div class="header-right">
            <h1>قـوات التحـالف الـعربي</h1>
            <h1>قـــوات درع الـــوطن</h1>
            <h1>القـائـــد العـــــام</h1>
        </div>
        
        <!-- Logo Section (Center) -->
        <div class="header-center">${logoSrc ? `<img src="${logoSrc}" alt="Logo">` : ''}</div>
        
        <!-- Date Box Section (Left in RTL) -->
        <div class="header-left">
            <div class="date-box">
                ${formData.showDate ? `
                <div class="header-row"><span class="header-label">التاريـخ:</span> <span>${dateStr}</span></div>
                <div class="header-row"><span class="header-label">اليـــوم:</span> <span>${dayName}</span></div>
                ` : ''}
                <div class="header-row"><span class="header-label">المرجع:</span> <span>_______</span></div>
                <div class="header-row"><span class="header-label">المرفقات:</span> <span>_______</span></div>
            </div>
        </div>
    </div>

    <!-- Padding for recipient to separate from header line -->
    <div style="height: 10px;"></div>

    <div class="recipient-section">
        <div class="recipient-name">${formData.to || ''}</div>
        <div class="recipient-name">${formData.to_the || 'المحترم'}</div>
    </div>

    <div class="greetings">${formData.greetings || 'السلام عليكم ورحمة الله وبركاته..'}</div>
    
    <div class="subject">${formData.subject_name || ''}</div>
    
    <div class="body-content">${formData.subject || ''}</div>
    
    <div class="ending">${formData.ending || 'وشكراً..'}</div>
    
    <div class="signature">${formData.sign || ''}</div>
    
    ${copyToItems ? `<div class="copy-to"><h4>نسخة إلى:</h4><ul>${copyToItems}</ul></div>` : ''}
</body>
</html>`;
}

// Convert to PDF using HTML template matching Word layout
ipcMain.handle('convert-to-pdf', async (event, { formData }) => {
    return new Promise(async (resolve, reject) => {
        try {
            const tempDir = os.tmpdir();
            const fullHtml = generateDocumentHtml(formData);
            const tempHtmlPath = path.join(tempDir, `autowriter_pdf_${Date.now()}.html`);
            fs.writeFileSync(tempHtmlPath, fullHtml, 'utf-8');

            // Create hidden window
            const hiddenWindow = new BrowserWindow({
                width: 794,
                height: 1123,
                show: false,
                webPreferences: {
                    nodeIntegration: false,
                    contextIsolation: true
                }
            });

            hiddenWindow.loadFile(tempHtmlPath);

            hiddenWindow.webContents.on('did-finish-load', async () => {
                try {
                    // Wait for rendering
                    await new Promise(r => setTimeout(r, 500));

                    // Generate PDF
                    const pdfBuffer = await hiddenWindow.webContents.printToPDF({
                        printBackground: true,
                        pageSize: 'A4',
                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                    });

                    // Cleanup
                    hiddenWindow.close();
                    try { fs.unlinkSync(tempHtmlPath); } catch (e) { }

                    resolve({ success: true, buffer: Array.from(pdfBuffer) });
                } catch (err) {
                    hiddenWindow.close();
                    try { fs.unlinkSync(tempHtmlPath); } catch (e) { }
                    reject(err);
                }
            });

            // Timeout
            setTimeout(() => {
                if (!hiddenWindow.isDestroyed()) {
                    hiddenWindow.close();
                }
                try { fs.unlinkSync(tempHtmlPath); } catch (e) { }
                reject(new Error('PDF conversion timeout'));
            }, 30000);

        } catch (err) {
            reject(err);
        }
    });
});

// Convert to Image using HTML template matching Word layout
ipcMain.handle('convert-to-image', async (event, { formData }) => {
    return new Promise(async (resolve, reject) => {
        try {
            const tempDir = os.tmpdir();
            const fullHtml = generateDocumentHtml(formData);
            const tempHtmlPath = path.join(tempDir, `autowriter_img_${Date.now()}.html`);
            fs.writeFileSync(tempHtmlPath, fullHtml, 'utf-8');

            // Create hidden window
            const hiddenWindow = new BrowserWindow({
                width: 794,
                height: 1123,
                show: false,
                webPreferences: {
                    nodeIntegration: false,
                    contextIsolation: true,
                    offscreen: true
                }
            });

            hiddenWindow.loadFile(tempHtmlPath);

            hiddenWindow.webContents.on('did-finish-load', async () => {
                try {
                    // Wait for rendering
                    await new Promise(r => setTimeout(r, 500));

                    // Capture screenshot
                    const image = await hiddenWindow.webContents.capturePage();
                    const pngBuffer = image.toPNG();

                    // Cleanup
                    hiddenWindow.close();
                    try { fs.unlinkSync(tempHtmlPath); } catch (e) { }

                    resolve({ success: true, buffer: Array.from(pngBuffer) });
                } catch (err) {
                    hiddenWindow.close();
                    try { fs.unlinkSync(tempHtmlPath); } catch (e) { }
                    reject(err);
                }
            });

            // Timeout
            setTimeout(() => {
                if (!hiddenWindow.isDestroyed()) {
                    hiddenWindow.close();
                }
                try { fs.unlinkSync(tempHtmlPath); } catch (e) { }
                reject(new Error('Image conversion timeout'));
            }, 30000);

        } catch (err) {
            reject(err);
        }
    });
});
