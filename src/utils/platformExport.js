/**
 * ===========================================
 * Platform Export Utility
 * ===========================================
 * 
 * Platform abstraction for document export:
 * - Desktop (Electron): Uses MS Word via PowerShell
 * - Android (Capacitor): Uses Filesystem/Share for saving files
 * 
 * @author AutoWriter
 * @version 2.0.0
 */

import { Capacitor, registerPlugin } from '@capacitor/core';
import { Filesystem, Directory } from '@capacitor/filesystem';
import { Share } from '@capacitor/share';
import { SplashScreen } from '@capacitor/splash-screen';

// Register LOK native plugin (for future use when LOK works)
const LOKPlugin = registerPlugin('LOKPlugin');

/**
 * Check if running on Android with native capabilities
 */
export const isAndroid = () => {
    return Capacitor.isNativePlatform() && Capacitor.getPlatform() === 'android';
};

/**
 * Check if running on Electron (desktop)
 */
export const isElectron = () => {
    return typeof window !== 'undefined' && window.electronAPI;
};

/**
 * Save file on Android using Filesystem plugin
 * @param {Uint8Array|number[]} buffer - File content
 * @param {string} filename - Name of file to save
 * @param {string} mimeType - MIME type of the file
 */
/**
 * Save file to Android cache directory (without sharing)
 */
export const saveToCacheAndroid = async (buffer, filename) => {
    // Convert buffer to base64
    const bytes = new Uint8Array(buffer);
    let binary = '';
    for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    const base64Data = btoa(binary);

    // Write file to cache directory
    const result = await Filesystem.writeFile({
        path: filename,
        data: base64Data,
        directory: Directory.Cache,
    });

    return { success: true, uri: result.uri };
};

/**
 * Share a file by URI
 */
export const shareFileAndroid = async (uri, filename, mimeType) => {
    try {
        await Share.share({
            title: filename,
            url: uri,
            dialogTitle: 'حفظ أو مشاركة الملف',
        });
        return { success: true };
    } catch (error) {
        if (error.message && (
            error.message.includes('cancel') ||
            error.message.includes('Cancel') ||
            error.message.includes('sharing is in progress')
        )) {
            return { success: true, cancelled: true };
        }
        return { success: false, error: error.message };
    }
};

/**
 * Save file on Android using Filesystem plugin
 * @param {Uint8Array|number[]} buffer - File content
 * @param {string} filename - Name of file to save
 * @param {string} mimeType - MIME type of the file
 */
export const saveFileAndroid = async (buffer, filename, mimeType) => {
    try {
        const saveResult = await saveToCacheAndroid(buffer, filename);
        console.log('File saved to:', saveResult.uri);

        return await shareFileAndroid(saveResult.uri, filename, mimeType);
    } catch (error) {
        console.error('saveFileAndroid error:', error);
        return { success: false, error: error.message };
    }
};

/**
 * Save multiple files on Android and share them all with one dialog
 * @param {Array<{buffer: Uint8Array, filename: string}>} files - Array of files to save
 * @param {string} dialogTitle - Title for the share dialog
 * @returns {Promise<{success: boolean, count?: number, error?: string}>}
 */
export const saveMultipleFilesAndroid = async (files, dialogTitle) => {
    const savedUris = [];
    try {
        for (const file of files) {
            // Convert buffer to base64
            const bytes = new Uint8Array(file.buffer);
            let binary = '';
            for (let i = 0; i < bytes.byteLength; i++) {
                binary += String.fromCharCode(bytes[i]);
            }
            const base64Data = btoa(binary);

            // Save to Cache directory
            const result = await Filesystem.writeFile({
                path: file.filename,
                data: base64Data,
                directory: Directory.Cache,
            });

            savedUris.push(result.uri);
            console.log('File saved to cache:', result.uri);
        }

        // Share all files with one dialog
        await Share.share({
            title: dialogTitle || 'مشاركة الصور',
            files: savedUris,
            dialogTitle: dialogTitle || 'حفظ أو مشاركة الصور',
        });

        return {
            success: true,
            count: savedUris.length
        };
    } catch (error) {
        console.error('saveMultipleFilesAndroid error:', error);
        // Check if user cancelled the share dialog or if there's a sharing-in-progress issue
        if (error.message && (
            error.message.includes('cancel') ||
            error.message.includes('Cancel') ||
            error.message.includes('sharing is in progress')
        )) {
            return { success: true, cancelled: true, count: savedUris.length };
        }
        return { success: false, error: error.message };
    }
};

/**
 * Convert DOCX buffer to PDF
 * NOTE: LOK currently crashes on Android due to asset bundling issues
 */
export const convertToPdf = async (docxBuffer) => {
    if (isElectron()) {
        console.log('Platform: Electron - Using MS Word for PDF conversion');
        return window.electronAPI.convertDocxToPdfWord(docxBuffer);
    }

    if (isAndroid()) {
        // Use LibreOfficeKit for PDF conversion
        console.log('Platform: Android - Using LibreOfficeKit for PDF conversion');
        try {
            const base64 = arrayToBase64(docxBuffer);
            const result = await LOKPlugin.convertToPdf({ docxBuffer: base64 });
            if (result.success && result.buffer) {
                const pdfBuffer = base64ToArray(result.buffer);
                return { success: true, buffer: pdfBuffer };
            } else {
                return { success: false, error: result.error || 'Conversion failed' };
            }
        } catch (error) {
            console.error('LOKPlugin convertToPdf error:', error);
            return { success: false, error: error.message };
        }
    }

    return { success: false, error: 'Platform not supported for PDF conversion' };
};

const sanitizeFileName = (name) => {
    return (name || 'image').replace(/[|&;$%@"<>()+,/\\:*?]/g, '').trim();
};

/**
 * Convert DOCX buffer to Image(s)
 */
export const convertToImage = async (docxBuffer, combinePages = true, subjectName = '') => {
    if (isElectron()) {
        console.log('Platform: Electron - Using MS Word + pdf-poppler for Image conversion');
        return window.electronAPI.convertDocxToImageWord(docxBuffer, combinePages, subjectName);
    }

    if (isAndroid()) {
        const safeName = sanitizeFileName(subjectName);
        console.log('Platform: Android - Using LibreOfficeKit for Image conversion, name:', safeName);

        // Extract extension to pass to plugin
        let extension = '.docx';
        if (subjectName && subjectName.lastIndexOf('.') !== -1) {
            extension = subjectName.substring(subjectName.lastIndexOf('.'));
        }

        try {
            const base64 = arrayToBase64(docxBuffer);

            if (combinePages) {
                // Single merged image
                const result = await LOKPlugin.convertToImage({
                    docxBuffer: base64,
                    extension: extension
                });
                if (result.success && result.buffer) {
                    const pngBuffer = base64ToArray(result.buffer);
                    return { success: true, buffer: pngBuffer };
                } else {
                    return { success: false, error: result.error || 'Image conversion failed' };
                }
            } else {
                // Separate pages
                const result = await LOKPlugin.convertToImages({
                    docxBuffer: base64,
                    baseName: safeName || 'page',
                    extension: extension
                });
                if (result.success && result.images && result.images.length > 0) {
                    // Convert each base64 image to buffer
                    const images = [];
                    for (let i = 0; i < result.images.length; i++) {
                        images.push(base64ToArray(result.images[i]));
                    }
                    return { success: true, images: images, count: images.length };
                } else {
                    return { success: false, error: result.error || 'Multi-page image conversion failed' };
                }
            }
        } catch (error) {
            console.error('LOKPlugin convertToImage error:', error);
            return { success: false, error: error.message };
        }
    }

    return { success: false, error: 'Platform not supported for Image conversion' };
};

/**
 * Check if the export functionality is ready
 */
export const isExportReady = async () => {
    if (isElectron()) {
        return true;
    }

    if (isAndroid()) {
        // For now, only DOCX export works on Android
        return true;
    }

    return false;
};

/**
 * Helper: Convert Uint8Array to Base64 string
 */
const arrayToBase64 = (array) => {
    let binary = '';
    const bytes = new Uint8Array(array);
    for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
};

/**
 * Helper: Convert Base64 string to Uint8Array
 */
const base64ToArray = (base64) => {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
};

/**
 * Clear all files from the cache directory (Android)
 * This should be called on app start and before exit to prevent storage bloat
 */
export const clearCacheAndroid = async () => {
    if (isAndroid()) {
        try {
            // List all files in cache directory
            const result = await Filesystem.readdir({
                path: '',
                directory: Directory.Cache
            });

            // Delete each file
            if (result.files && result.files.length > 0) {
                for (const file of result.files) {
                    try {
                        // file can be a string (older API) or an object with name property
                        const fileName = typeof file === 'string' ? file : file.name;
                        await Filesystem.deleteFile({
                            path: fileName,
                            directory: Directory.Cache
                        });
                        console.log('Deleted cached file:', fileName);
                    } catch (deleteError) {
                        console.warn('Failed to delete file:', file, deleteError);
                    }
                }
                console.log(`Cleared ${result.files.length} cached files`);
            }
            return { success: true, count: result.files?.length || 0 };
        } catch (e) {
            // If directory doesn't exist or is empty, that's fine
            console.log('Cache clear info:', e.message);
            return { success: true, count: 0 };
        }
    }
    return { success: false, error: 'Not on Android' };
};

/**
 * Convert document to image FILES (returns paths, not Base64)
 * Uses file-path based processing to avoid memory issues with large PDFs
 * @param {Uint8Array} docxBuffer - Document buffer
 * @param {string} subjectName - Base name for output files
 * @param {function} onProgress - Callback(current, total, percent)
 * @returns {Promise<{success: boolean, paths?: string[], count?: number, error?: string}>}
 */
export const convertToImageFiles = async (docxBuffer, subjectName = '', onProgress = null) => {
    if (!isAndroid()) {
        return { success: false, error: 'Only supported on Android' };
    }

    const safeName = sanitizeFileName(subjectName) || 'page';
    let extension = '.docx';
    if (subjectName && subjectName.lastIndexOf('.') !== -1) {
        extension = subjectName.substring(subjectName.lastIndexOf('.'));
    }

    try {
        // Register progress listener if callback provided
        let progressListener = null;
        if (onProgress) {
            progressListener = await LOKPlugin.addListener('conversionProgress', (data) => {
                onProgress(data.current, data.total, data.percent);
            });
        }

        const base64 = arrayToBase64(docxBuffer);
        const result = await LOKPlugin.convertToImageFiles({
            docxBuffer: base64,
            baseName: safeName,
            extension: extension
        });

        // Remove listener
        if (progressListener) {
            progressListener.remove();
        }

        if (result.success && result.paths) {
            return {
                success: true,
                paths: Array.isArray(result.paths) ? result.paths : [],
                count: result.count
            };
        } else {
            return { success: false, error: result.error || 'Conversion failed' };
        }
    } catch (error) {
        console.error('convertToImageFiles error:', error);
        return { success: false, error: error.message };
    }
};

/**
 * Get single image as Base64 from file path
 * @param {string} path - File path to image
 * @returns {Promise<string|null>} Base64 string or null
 */
export const getImageBase64ByPath = async (path) => {
    if (!isAndroid()) return null;

    try {
        const result = await LOKPlugin.getImageAsBase64({ path });
        return result.success ? result.data : null;
    } catch (error) {
        console.error('getImageBase64ByPath error:', error);
        return null;
    }
};

/**
 * Delete temporary image files
 * @param {string[]} paths - Array of file paths to delete
 */
export const cleanupImageFiles = async (paths) => {
    if (!isAndroid() || !paths || paths.length === 0) return;

    try {
        await LOKPlugin.deleteImageFiles({ paths });
    } catch (error) {
        console.error('cleanupImageFiles error:', error);
    }
};

export default {
    convertToPdf,
    convertToImage,
    convertToImageFiles,
    getImageBase64ByPath,
    cleanupImageFiles,
    saveFileAndroid,
    saveToCacheAndroid,
    shareFileAndroid,
    clearCacheAndroid,
    isExportReady,
    isElectron,
    isAndroid
};

/**
 * Check if the app was launched with a shared image (or received one while running)
 * Returns array of base64 strings or null
 */
export const checkForSharedImage = async () => {
    if (isAndroid()) {
        try {
            console.log('Checking for shared image...');
            const result = await LOKPlugin.getSharedImage();
            if (result.hasImage) {
                console.log('Shared content found');
                if (result.files && Array.isArray(result.files)) {
                    return result.files;
                } else if (result.images && Array.isArray(result.images)) {
                    return result.images;
                } else if (result.image) {
                    return [result.image];
                }
            }
        } catch (e) {
            console.error('Error checking shared image', e);
        }
    }
    return null;
};

/**
 * Dynamic Import Wrapper for Native Plugins
 */

// Pick Images (Android)
export const pickImagesAndroid = async (multiple = true) => {
    if (isAndroid()) {
        try {
            const { FilePicker } = await import('@capawesome/capacitor-file-picker');
            const result = await FilePicker.pickMedia({
                multiple: multiple,
                readData: true
            });
            return result.files;
        } catch (e) {
            console.error('FilePicker error:', e);
            throw e;
        }
    }
    return [];
};

// Pick Files (Android) - supports images, Word docs, Excel, PowerPoint, and PDFs
// Two-step process: pick without data → filter by extension → read only supported files
// Returns { files: [], skippedCount: number }
export const pickFilesAndroid = async (multiple = true) => {
    if (isAndroid()) {
        try {
            const { FilePicker } = await import('@capawesome/capacitor-file-picker');

            // Supported extensions
            const supportedDocs = ['.docx', '.doc', '.pptx', '.ppt', '.xlsx', '.xls', '.pdf'];
            const supportedImages = ['.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp'];

            // Step 1: Pick files WITHOUT reading data (fast - only gets metadata)
            const result = await FilePicker.pickFiles({
                multiple: multiple,
                readData: false,
                types: [
                    'image/*',
                    'application/pdf',
                    'application/msword',
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    'application/vnd.ms-powerpoint',
                    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                    'application/vnd.ms-excel',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                ]
            });

            // Step 2: Filter by extension (skip unsupported like videos)
            const supportedFiles = [];
            let skippedCount = 0;

            for (const file of result.files) {
                const name = (file.name || '').toLowerCase();
                const isSupported = supportedDocs.some(ext => name.endsWith(ext)) ||
                    supportedImages.some(ext => name.endsWith(ext));

                if (isSupported) {
                    supportedFiles.push(file);
                } else {
                    skippedCount++;
                }
            }

            // Step 3: Read data ONLY for supported files
            const { Filesystem } = await import('@capacitor/filesystem');
            const filesWithData = [];

            for (const file of supportedFiles) {
                try {
                    const fileData = await Filesystem.readFile({ path: file.path });
                    filesWithData.push({
                        ...file,
                        data: fileData.data
                    });
                } catch (readError) {
                    console.error('Error reading file:', file.name, readError);
                }
            }

            return { files: filesWithData, skippedCount };
        } catch (e) {
            console.error('FilePicker pickFiles error:', e);
            throw e;
        }
    }
    return { files: [], skippedCount: 0 };
};

// Open Browser (Android)
export const openBrowserAndroid = async (url) => {
    if (isAndroid()) {
        try {
            const { Browser } = await import('@capacitor/browser');
            await Browser.open({ url });
            return;
        } catch (e) {
            console.error('Browser error:', e);
        }
    }
    // Fallback?
};

// Add App Listener (Android)
export const addAppListenerAndroid = async (eventName, callback) => {
    if (isAndroid()) {
        try {
            const { App: CapacitorApp } = await import('@capacitor/app');
            return await CapacitorApp.addListener(eventName, callback);
        } catch (e) {
            console.error('App listener error:', e);
        }
    }
    return { remove: () => { } };
};

export const removeAllAppListenersAndroid = async () => {
    if (isAndroid()) {
        try {
            const { App: CapacitorApp } = await import('@capacitor/app');
            await CapacitorApp.removeAllListeners();
        } catch (e) { }
    }
};

// Exit App (Android)
export const exitAppAndroid = async () => {
    if (isAndroid()) {
        try {
            // Clear cache before exiting
            await clearCacheAndroid();

            const { App: CapacitorApp } = await import('@capacitor/app');
            await CapacitorApp.exitApp();
        } catch (e) {
            console.error('Exit app error:', e);
        }
    }
};

// Hide Native Splash Screen (Android)
export const hideSplashScreen = async () => {
    if (isAndroid()) {
        try {
            await SplashScreen.hide();
        } catch (e) {
            console.error('Hide splash screen error:', e);
        }
    }
};

