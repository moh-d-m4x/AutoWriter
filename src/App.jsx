import { useState, useEffect, useRef } from 'react'
import PizZip from 'pizzip'
import {
    isElectron, isAndroid, convertToPdf, convertToImage, convertToImageFiles,
    getImageBase64ByPath, cleanupImageFiles,
    saveFileAndroid, saveToCacheAndroid, shareFileAndroid, saveMultipleFilesAndroid,
    checkForSharedImage, pickImagesAndroid, pickFilesAndroid, openBrowserAndroid, addAppListenerAndroid,
    exitAppAndroid, clearCacheAndroid
} from './utils/platformExport'
import { jsPDF } from "jspdf"


const compressImage = (dataUrl, quality = 0.8) => {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
            const canvas = document.createElement('canvas');
            let width = img.width;
            let height = img.height;

            // Cap dimensions to reasonable size (e.g. A4 @ 300dpi approx 2480px width)
            const MAX_WIDTH = 2480;
            const MAX_HEIGHT = 3508;

            if (width > MAX_WIDTH || height > MAX_HEIGHT) {
                const ratio = width / height;
                if (width > MAX_WIDTH) {
                    width = MAX_WIDTH;
                    height = width / ratio;
                }
                if (height > MAX_HEIGHT) {
                    height = MAX_HEIGHT;
                    width = height * ratio;
                }
            }

            canvas.width = width;
            canvas.height = height;

            const ctx = canvas.getContext('2d');
            // White background for transparent PNGs converted to JPEG
            ctx.fillStyle = '#FFFFFF';
            ctx.fillRect(0, 0, width, height);
            ctx.drawImage(img, 0, 0, width, height);
            resolve(canvas.toDataURL('image/jpeg', quality));
        };
        img.onerror = reject;
        img.src = dataUrl;
    });
};

function App() {

    // Toast notification state
    const [toast, setToast] = useState({ show: false, message: '', type: 'success' });

    // Export loading state
    const [isExporting, setIsExporting] = useState(false);
    const [exportFormat, setExportFormat] = useState(null); // Track current export format
    const [exportProgress, setExportProgress] = useState(0); // Progress percentage (0-100)
    const [fileProgress, setFileProgress] = useState({ current: 0, total: 0 }); // Multi-file progress (1/3, 2/3)

    // Export cancellation ref (mutable, survives re-renders)
    const exportCancelledRef = useRef(false);

    // Temp image paths for cleanup (file-path based processing)
    const tempImagePathsRef = useRef([]);

    // Combine image pages option (default: false = separate images)
    const [combineImagePages, setCombineImagePages] = useState(false);

    // Confirmation modal state
    const [confirmModal, setConfirmModal] = useState({
        show: false,
        message: '',
        onConfirm: null
    });

    // Preview modal state
    const [previewState, setPreviewState] = useState({
        show: false,
        images: [],        // Array of base64 data URLs
        currentPage: 0,
        zoom: 1.0,
        panX: 0,           // Pan offset X
        panY: 0,           // Pan offset Y
        isLoading: false,
        isExternal: false
    });

    // Added images (to be appended to exports)
    const [addedImages, setAddedImages] = useState([]);

    // Drag-and-drop state (PC only)
    const [isDragging, setIsDragging] = useState(false);
    const dragCounter = useRef(0);

    // Source selection dialog state (Android only - camera or file picker)
    const [showSourceDialog, setShowSourceDialog] = useState(false);

    // Show toast notification
    const showToast = (message, type = 'success') => {
        setToast({ show: true, message, type });
        setTimeout(() => {
            setToast({ show: false, message: '', type: 'success' });
        }, 3000);
    };

    // Form state with default values
    const [formData, setFormData] = useState({
        logo: null,
        logoBase64: '',
        logoName: '',
        parent_company: '',
        subsidiary_company: '',
        from: '',
        to: 'الأخ / ',
        to_the: 'المحترم',
        greetings: '',
        subject_name: 'الموضوع / ',
        subject: 'في البَدْء نهديكم أطيب التحايا متمنين لكم النجاح في مهامكم العملية وإشارة إلى الموضوع أعلاه،‏ ',
        ending: '',
        sign: '',
        copy_to: '',
        showDate: false,
        useTable: false,
        tableData: [['م', 'الاسم', 'ملاحظة'], ['1', '', '']]
    });

    // Persistent fields (to_the is NOT persistent - resets to default each session)
    const persistentFields = ['logo', 'logoBase64', 'logoName', 'parent_company', 'subsidiary_company', 'from', 'greetings', 'ending', 'sign', 'copy_to'];

    // Track if we've loaded data
    const [isLoaded, setIsLoaded] = useState(false);

    // Load persistent data on mount
    useEffect(() => {
        const saved = localStorage.getItem('autowriter_data');
        if (saved) {
            try {
                const parsed = JSON.parse(saved);
                setFormData(prev => ({ ...prev, ...parsed }));
            } catch (e) {
                console.error('Error loading saved data:', e);
            }
        }
        setIsLoaded(true);
    }, []);

    // Save persistent data on change (only after initial load)
    useEffect(() => {
        if (!isLoaded) return; // Don't save until we've loaded

        const toSave = {};
        persistentFields.forEach(field => {
            toSave[field] = formData[field];
        });
        localStorage.setItem('autowriter_data', JSON.stringify(toSave));
    }, [formData, isLoaded]);

    // Clear cache on app startup (Android only)
    // This removes any leftover temporary files from previous sessions
    useEffect(() => {
        if (isAndroid()) {
            clearCacheAndroid().then(result => {
                if (result.count > 0) {
                    console.log(`Startup: Cleared ${result.count} cached files from previous session`);
                }
            }).catch(e => {
                console.warn('Startup cache clear failed:', e);
            });
        }
    }, []);

    // Auto-resize all textareas on initial load
    useEffect(() => {
        if (isLoaded) {
            // Use setTimeout to ensure DOM is ready
            setTimeout(() => {
                const textareas = document.querySelectorAll('textarea');
                textareas.forEach(textarea => {
                    textarea.style.height = 'auto';
                    textarea.style.height = textarea.scrollHeight + 'px';
                });
            }, 100);
        }
    }, [isLoaded]);

    // Disable body scroll when preview is open
    useEffect(() => {
        if (previewState.show) {
            document.body.style.overflow = 'hidden';
        } else {
            document.body.style.overflow = '';
        }
        return () => {
            document.body.style.overflow = '';
        };
    }, [previewState.show]);

    // Handle files opened via "Open with" on PC (Electron only)
    useEffect(() => {
        if (!isElectron() || !window.electronAPI?.onFileOpened) return;

        const handleFileOpened = async (fileData) => {
            console.log('File opened via "Open with":', fileData.name);

            const name = fileData.name.toLowerCase();
            const finalImages = [];
            let hasConverted = false;

            setIsExporting(true);
            setExportFormat('docx-conversion');

            try {
                // Convert base64 to Uint8Array
                const binaryString = window.atob(fileData.data);
                const bytes = new Uint8Array(binaryString.length);
                for (let i = 0; i < binaryString.length; i++) {
                    bytes[i] = binaryString.charCodeAt(i);
                }

                // Handle Word Documents, PowerPoint, and PDFs
                if (name.endsWith('.docx') || name.endsWith('.doc') || name.endsWith('.pptx') || name.endsWith('.ppt') || name.endsWith('.pdf')) {
                    hasConverted = true;

                    const result = await window.electronAPI.convertDocxToImageWord(
                        Array.from(bytes),
                        false,
                        fileData.name
                    );

                    if (result.success && result.images) {
                        for (const imgData of result.images) {
                            const imgBuffer = imgData.buffer || imgData;
                            let binary = '';
                            for (let i = 0; i < imgBuffer.length; i++) {
                                binary += String.fromCharCode(imgBuffer[i]);
                            }
                            finalImages.push(`data:image/png;base64,${btoa(binary)}`);
                        }
                    } else {
                        showToast(`فشل تحويل: ${fileData.name}`, 'error');
                    }
                } else {
                    // Handle images directly
                    let mime = 'image/png';
                    if (name.endsWith('.jpg') || name.endsWith('.jpeg')) mime = 'image/jpeg';
                    if (name.endsWith('.webp')) mime = 'image/webp';
                    if (name.endsWith('.gif')) mime = 'image/gif';
                    finalImages.push(`data:${mime};base64,${fileData.data}`);
                }

                if (finalImages.length > 0) {
                    // If preview is already open, append; otherwise open new preview
                    if (previewState.show) {
                        setPreviewState(prev => ({
                            ...prev,
                            images: [...prev.images, ...finalImages]
                        }));
                        setAddedImages(prev => [...prev, ...finalImages]);
                        showToast(`تم إضافة ${finalImages.length} ${finalImages.length > 1 ? 'صفحات' : 'صفحة'} بنجاح`);
                    } else {
                        setPreviewState({
                            show: true,
                            images: finalImages,
                            currentPage: 0,
                            zoom: 1.0,
                            panX: 0,
                            panY: 0,
                            isLoading: false,
                            isExternal: true
                        });
                        if (hasConverted) {
                            showToast('تم تحويل الملف بنجاح', 'success');
                        } else {
                            showToast('تم فتح الصورة', 'success');
                        }
                    }
                }
            } catch (error) {
                console.error('Error processing opened file:', error);
                showToast('خطأ في فتح الملف', 'error');
            } finally {
                setIsExporting(false);
            }
        };

        const cleanup = window.electronAPI.onFileOpened(handleFileOpened);
        return cleanup;
    }, [previewState.show]);

    // Check for shared images on load and when receiving event
    const isProcessingSharedRef = useRef(false);

    useEffect(() => {
        const handleSharedImage = async () => {
            // Prevent duplicate concurrent calls
            if (isProcessingSharedRef.current) {
                console.log('handleSharedImage: Already processing, skipping');
                return;
            }
            isProcessingSharedRef.current = true;

            try {
                // Show loading immediately while fetching shared file
                setIsExporting(true);
                setExportFormat('docx-conversion');
                setExportProgress(0);

                const sharedItems = await checkForSharedImage();

                // Early exit if no shared items
                if (!sharedItems || sharedItems.length === 0) {
                    setIsExporting(false);
                    isProcessingSharedRef.current = false;
                    return;
                }

                if (sharedItems.length > 0) {
                    const finalImages = [];
                    let hasConverted = false;
                    let skippedCount = 0;

                    // Supported file extensions
                    const supportedDocs = ['.docx', '.doc', '.pptx', '.ppt', '.xlsx', '.xls', '.pdf'];
                    const supportedImages = ['.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp'];

                    // Count document files for progress tracking
                    const docFiles = sharedItems.filter(item => {
                        if (typeof item === 'string') return false;
                        const name = (item.name || '').toLowerCase();
                        return supportedDocs.some(ext => name.endsWith(ext));
                    });
                    const totalDocs = docFiles.length;
                    let currentDoc = 0;

                    for (const item of sharedItems) {
                        // Legacy string format
                        if (typeof item === 'string') {
                            finalImages.push(`data:image/png;base64,${item}`);
                        }
                        // New object format
                        else if (item.name) {
                            const name = item.name.toLowerCase();
                            const isDoc = supportedDocs.some(ext => name.endsWith(ext));
                            const isImage = supportedImages.some(ext => name.endsWith(ext));

                            // Skip unsupported files
                            if (!isDoc && !isImage) {
                                skippedCount++;
                                continue;
                            }

                            // Handle documents (PDF, Word, PowerPoint, Excel)
                            // Streaming approach: native processes file directly from path
                            if (item.isDocument && item.path) {
                                currentDoc++;
                                hasConverted = true;
                                setIsExporting(true);
                                setExportFormat('docx-conversion');
                                setExportProgress(0);
                                setFileProgress({ current: currentDoc, total: totalDocs });

                                // Import the streaming converter
                                const { convertDocumentFromPath } = await import('./utils/platformExport');

                                // Convert directly from file path - NO memory loading!
                                const result = await convertDocumentFromPath(item.path, item.name, (current, total, percent) => {
                                    setExportProgress(percent);
                                });

                                if (result.success && result.paths && result.paths.length > 0) {
                                    tempImagePathsRef.current = [...tempImagePathsRef.current, ...result.paths];

                                    // Load images one by one from file paths
                                    for (const path of result.paths) {
                                        const base64 = await getImageBase64ByPath(path);
                                        if (base64) {
                                            finalImages.push(`data:image/png;base64,${base64}`);
                                        }
                                    }
                                } else {
                                    showToast(result.error || 'فشل تحويل الملف', 'error');
                                }
                                setIsExporting(false);
                                setExportProgress(0);
                                setFileProgress({ current: 0, total: 0 });
                            }
                            // Legacy approach: item.data contains Base64 (for backwards compatibility)
                            else if (isDoc && item.data) {
                                currentDoc++;
                                hasConverted = true;
                                setIsExporting(true);
                                setExportFormat('docx-conversion');
                                setExportProgress(0);
                                setFileProgress({ current: currentDoc, total: totalDocs });

                                // Convert Base64 to Uint8Array
                                const binaryString = window.atob(item.data);
                                const bytes = new Uint8Array(binaryString.length);
                                for (let i = 0; i < binaryString.length; i++) {
                                    bytes[i] = binaryString.charCodeAt(i);
                                }

                                // Use new file-path based conversion with progress
                                const result = await convertToImageFiles(bytes, item.name, (current, total, percent) => {
                                    setExportProgress(percent);
                                });

                                if (result.success && result.paths && result.paths.length > 0) {
                                    // Store paths for cleanup later
                                    tempImagePathsRef.current = result.paths;

                                    // Load images one by one from file paths
                                    for (const path of result.paths) {
                                        const base64 = await getImageBase64ByPath(path);
                                        if (base64) {
                                            finalImages.push(`data:image/png;base64,${base64}`);
                                        }
                                    }
                                } else {
                                    showToast(result.error || 'فشل تحويل الملف', 'error');
                                }
                                setIsExporting(false);
                                setExportProgress(0);
                                setFileProgress({ current: 0, total: 0 });
                            } else if (isImage && item.data) {
                                // Supported Image
                                let mime = 'image/png';
                                if (name.endsWith('.jpg') || name.endsWith('.jpeg')) mime = 'image/jpeg';
                                if (name.endsWith('.webp')) mime = 'image/webp';
                                if (name.endsWith('.gif')) mime = 'image/gif';
                                if (name.endsWith('.bmp')) mime = 'image/bmp';
                                finalImages.push(`data:${mime};base64,${item.data}`);
                            }
                        }
                    }

                    // Show notification about skipped unsupported files
                    if (skippedCount > 0) {
                        showToast(`تم تجاهل ${skippedCount} ${skippedCount > 1 ? 'ملفات غير مدعومة' : 'ملف غير مدعوم'}`, 'error');
                    }

                    if (finalImages.length > 0) {
                        setPreviewState({
                            show: true,
                            images: finalImages,
                            currentPage: 0,
                            zoom: 1.0,
                            panX: 0,
                            panY: 0,
                            isLoading: false,
                            isExternal: true
                        });
                        if (hasConverted) showToast('تم تحويل الملف بنجاح', 'success');
                        else showToast(`تم استلام ${finalImages.length > 1 ? finalImages.length + ' صور' : 'صورة'}`, 'success');

                        // Cleanup temp image files after loading into memory
                        if (tempImagePathsRef.current.length > 0) {
                            cleanupImageFiles(tempImagePathsRef.current);
                            tempImagePathsRef.current = [];
                        }
                    }
                }
            } catch (e) {
                console.error("Error handling shared file:", e);
                showToast('خطأ في استلام الملف', 'error');
            } finally {
                isProcessingSharedRef.current = false;
                setIsExporting(false);
                setExportProgress(0);
                setFileProgress({ current: 0, total: 0 });
            }
        };

        if (isAndroid()) {
            // Check immediately on mount
            handleSharedImage();

            // Listen for event from native code
            window.addEventListener('appSharedImageAvailable', handleSharedImage);

            // Also check on app resume
            addAppListenerAndroid('resume', handleSharedImage);
        }

        return () => {
            window.removeEventListener('appSharedImageAvailable', handleSharedImage);
        };
    }, []);

    // Handle Android hardware back button
    useEffect(() => {
        let backListener;
        const setupBackListener = async () => {
            if (isAndroid()) {
                backListener = await addAppListenerAndroid('backButton', ({ canGoBack }) => {
                    // Priority 0: Close Confirmation Modal
                    if (confirmModal.show) {
                        setConfirmModal(prev => ({ ...prev, show: false }));
                        return;
                    }

                    // Priority 1: Cancel Exporting
                    if (isExporting) {
                        // Exclude Word export from cancellation (no cancel button)
                        if (exportFormat === 'docx') {
                            return;
                        }

                        exportCancelledRef.current = true;
                        setIsExporting(false);
                        return;
                    }

                    // Priority 2: Close Preview
                    if (previewState.show) {
                        setPreviewState(prev => ({
                            ...prev,
                            show: false, images: [], currentPage: 0, zoom: 1.0, panX: 0, panY: 0, isLoading: false
                        }));
                        setAddedImages([]); // Clear added images when closing preview
                        return;
                    }

                    // Priority 3: Default (Exit -> Confirm Exit)
                    if (!canGoBack) {
                        setConfirmModal({
                            show: true,
                            message: 'هل أنت متأكد من الخروج من التطبيق؟',
                            onConfirm: () => {
                                exitAppAndroid();
                            }
                        });
                    }
                });
            }
        };
        setupBackListener();
        return () => {
            if (backListener) backListener.remove();
        };
    }, [previewState.show, isExporting, exportFormat, confirmModal.show]);

    // Auto-resize textarea based on content
    const autoResize = (element) => {
        if (element && element.tagName === 'TEXTAREA') {
            element.style.height = 'auto';
            element.style.height = element.scrollHeight + 'px';
        }
    };

    // Handle input changes
    const handleChange = (e) => {
        const { name, value, type, checked } = e.target;
        setFormData(prev => ({
            ...prev,
            [name]: type === 'checkbox' ? checked : value
        }));
        // Auto-resize textareas
        autoResize(e.target);
    };

    // Handle logo selection
    const handleLogoSelect = async () => {
        // Supported image extensions
        const supportedImages = ['.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp'];

        if (window.electronAPI) {
            // Electron (Desktop)
            const result = await window.electronAPI.selectImage();
            if (result) {
                // Validate file type
                const ext = result.name ? '.' + result.name.toLowerCase().split('.').pop() : '';
                if (!supportedImages.includes(ext)) {
                    showToast('يرجى اختيار ملف صورة فقط (PNG, JPG, WEBP)', 'error');
                    return;
                }
                setFormData(prev => ({
                    ...prev,
                    logo: result.path,
                    logoBase64: result.base64,
                    logoName: result.name
                }));
            }
        } else if (isAndroid()) {
            // Android (Capacitor) - use FilePicker to preserve original format (PNG transparency)
            try {
                const files = await pickImagesAndroid(false);
                const result = { files };

                if (result && result.files && result.files.length > 0) {
                    const file = result.files[0];

                    // Validate file type
                    const name = (file.name || '').toLowerCase();
                    const isImage = supportedImages.some(ext => name.endsWith(ext));

                    if (!isImage) {
                        showToast('يرجى اختيار ملف صورة فقط (PNG, JPG, WEBP)', 'error');
                        return;
                    }

                    if (file.data) {
                        // Determine format from MIME type
                        const mimeType = file.mimeType || 'image/png';
                        const format = mimeType.includes('png') ? 'png' :
                            mimeType.includes('gif') ? 'gif' :
                                mimeType.includes('webp') ? 'webp' : 'jpeg';

                        setFormData(prev => ({
                            ...prev,
                            logo: 'android-selected',
                            logoBase64: file.data,
                            logoName: file.name || `image.${format}`
                        }));
                    }
                }
            } catch (error) {
                console.error('Error picking image:', error);
                // User cancelled or error - do nothing
            }
        }
    };

    // Handle logo removal
    const handleLogoClear = (e) => {
        e.stopPropagation(); // Prevent triggering logo select
        setFormData(prev => ({
            ...prev,
            logo: null,
            logoBase64: '',
            logoName: ''
        }));
    };

    // External Add Files (Images, Word, PDF)
    const handleExternalAddImages = async () => {
        // Android: Show source selection dialog
        if (isAndroid()) {
            setShowSourceDialog(true);
            return;
        }

        // PC: Use Electron file dialog for all supported types
        if (isElectron() && window.electronAPI) {
            try {
                let newImages = [];
                const result = await window.electronAPI.selectFiles();
                if (result && result.length > 0) {
                    let hasConverted = false;
                    setIsExporting(true);
                    setExportFormat('docx-conversion');

                    try {
                        for (const file of result) {
                            const name = file.name.toLowerCase();

                            // Handle Word Documents, PowerPoint, Excel, and PDFs
                            if (name.endsWith('.docx') || name.endsWith('.doc') || name.endsWith('.pptx') || name.endsWith('.ppt') || name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.pdf')) {
                                hasConverted = true;

                                // Convert Base64 to Uint8Array
                                const binaryString = window.atob(file.base64);
                                const bytes = new Uint8Array(binaryString.length);
                                for (let i = 0; i < binaryString.length; i++) {
                                    bytes[i] = binaryString.charCodeAt(i);
                                }

                                // Convert to Images using Electron API
                                const convResult = await window.electronAPI.convertDocxToImageWord(
                                    Array.from(bytes),
                                    false,
                                    file.name
                                );

                                if (convResult.success && convResult.images) {
                                    for (const imgData of convResult.images) {
                                        const imgBuffer = imgData.buffer || imgData;
                                        let binary = '';
                                        for (let i = 0; i < imgBuffer.length; i++) {
                                            binary += String.fromCharCode(imgBuffer[i]);
                                        }
                                        newImages.push(`data:image/png;base64,${btoa(binary)}`);
                                    }
                                } else {
                                    showToast(`فشل تحويل: ${file.name}`, 'error');
                                }
                            } else {
                                // Assume Image - use proper mime type
                                let mime = 'image/png';
                                if (name.endsWith('.jpg') || name.endsWith('.jpeg')) mime = 'image/jpeg';
                                if (name.endsWith('.webp')) mime = 'image/webp';
                                if (name.endsWith('.gif')) mime = 'image/gif';
                                if (name.endsWith('.bmp')) mime = 'image/bmp';
                                newImages.push(`data:${mime};base64,${file.base64}`);
                            }
                        }

                        if (hasConverted && newImages.length > 0) {
                            showToast('تم تحويل المستندات بنجاح', 'success');
                        }
                    } finally {
                        setIsExporting(false);
                    }
                }

                if (newImages.length > 0) {
                    setPreviewState(prev => ({ ...prev, images: [...prev.images, ...newImages] }));
                    setAddedImages(prev => [...prev, ...newImages]);
                    showToast(`تم إضافة ${newImages.length} ${newImages.length > 1 ? 'صفحات' : 'صفحة'} بنجاح`);
                }
            } catch (e) { console.error(e); }
        }
    };

    // Android: Add from Camera
    const handleAndroidAddFromCamera = async () => {
        setShowSourceDialog(false);
        try {
            const { Camera, CameraResultType, CameraSource } = await import('@capacitor/camera');
            const photo = await Camera.getPhoto({
                quality: 90,
                allowEditing: false,
                resultType: CameraResultType.Base64,
                source: CameraSource.Camera
            });

            if (photo.base64String) {
                const dataUrl = `data:image/${photo.format || 'jpeg'};base64,${photo.base64String}`;
                setPreviewState(prev => ({ ...prev, images: [...prev.images, dataUrl] }));
                setAddedImages(prev => [...prev, dataUrl]);
                showToast('تم التقاط الصورة بنجاح');
            }
        } catch (e) {
            if (!e.message?.includes('cancel')) {
                console.error('Camera error:', e);
                showToast('خطأ في الكاميرا', 'error');
            }
        }
    };

    // Android: Add from File Picker
    const handleAndroidAddFromFilePicker = async () => {
        setShowSourceDialog(false);
        try {
            let newImages = [];
            // pickFilesAndroid now returns { files, skippedCount }
            const result = await pickFilesAndroid(true);
            const files = result.files || [];
            const skippedCount = result.skippedCount || 0;

            // Show notification about skipped unsupported files
            if (skippedCount > 0) {
                showToast(`تم تجاهل ${skippedCount} ${skippedCount > 1 ? 'ملفات غير مدعومة' : 'ملف غير مدعوم'}`, 'error');
            }

            if (files && files.length > 0) {
                let hasConverted = false;

                // Supported file extensions
                const supportedDocs = ['.docx', '.doc', '.pptx', '.ppt', '.xlsx', '.xls', '.pdf'];

                // Count document files for progress tracking
                const docFiles = files.filter(file => {
                    const name = (file.name || '').toLowerCase();
                    return supportedDocs.some(ext => name.endsWith(ext));
                });
                const totalDocs = docFiles.length;
                let currentDoc = 0;

                for (const file of files) {
                    const name = (file.name || '').toLowerCase();
                    const isDoc = supportedDocs.some(ext => name.endsWith(ext));

                    // Handle Word Documents, PowerPoint, Excel, and PDF
                    if (isDoc) {
                        currentDoc++;
                        hasConverted = true;
                        setIsExporting(true);
                        setExportFormat('docx-conversion');
                        setExportProgress(0);
                        setFileProgress({ current: currentDoc, total: totalDocs });

                        // Convert Base64 to Uint8Array
                        const binaryString = window.atob(file.data);
                        const bytes = new Uint8Array(binaryString.length);
                        for (let i = 0; i < binaryString.length; i++) {
                            bytes[i] = binaryString.charCodeAt(i);
                        }

                        // Use file-path based conversion with progress
                        const convResult = await convertToImageFiles(bytes, file.name, (current, total, percent) => {
                            setExportProgress(percent);
                        });

                        if (convResult.success && convResult.paths && convResult.paths.length > 0) {
                            // Load images from file paths
                            for (const path of convResult.paths) {
                                const base64 = await getImageBase64ByPath(path);
                                if (base64) {
                                    newImages.push(`data:image/png;base64,${base64}`);
                                }
                            }
                            // Cleanup temp files
                            cleanupImageFiles(convResult.paths);
                        } else {
                            showToast(convResult.error || 'فشل تحويل الملف', 'error');
                        }
                        setIsExporting(false);
                        setExportProgress(0);
                        setFileProgress({ current: 0, total: 0 });
                    } else {
                        // Image file
                        const mime = file.mimeType || 'image/png';
                        newImages.push(`data:${mime};base64,${file.data}`);
                    }
                }

                if (hasConverted && newImages.length > 0) {
                    showToast('تم تحويل المستندات بنجاح', 'success');
                }
            }

            if (newImages.length > 0) {
                setPreviewState(prev => ({ ...prev, images: [...prev.images, ...newImages] }));
                setAddedImages(prev => [...prev, ...newImages]);
                showToast(`تم إضافة ${newImages.length} ${newImages.length > 1 ? 'صفحات' : 'صفحة'} بنجاح`);
            }
        } catch (e) { console.error(e); }
    };

    // Delete current image from preview
    const handleDeleteCurrentImage = () => {
        setPreviewState(prev => {
            const newImages = [...prev.images];
            const deletedImage = newImages.splice(prev.currentPage, 1)[0];

            // Also remove from addedImages if it was an added image
            setAddedImages(prevAdded => prevAdded.filter(img => img !== deletedImage));

            // Adjust current page if needed
            let newPage = prev.currentPage;
            if (newPage >= newImages.length && newImages.length > 0) {
                newPage = newImages.length - 1;
            }

            // If no images left, close preview
            if (newImages.length === 0) {
                showToast('تم حذف جميع الصور');
                return { ...prev, show: false, images: [], currentPage: 0 };
            }

            showToast('تم حذف الصورة');
            return { ...prev, images: newImages, currentPage: newPage };
        });
    };

    // Sanitize filename - replace Windows unsupported characters with dash
    const sanitizeFileName = (name) => {
        if (!name) return 'document';
        return name.replace(/[\\/:*?"<>|]/g, '-');
    };

    // Open URL in system's default browser
    const openExternalUrl = async (url) => {
        try {
            if (isElectron()) {
                // Use Electron's shell.openExternal
                await window.electronAPI.openExternal(url);
            } else if (isAndroid()) {
                // Use Capacitor Browser plugin
                await openBrowserAndroid(url);
            } else {
                // Fallback for web
                window.open(url, '_blank');
            }
        } catch (error) {
            console.error('Failed to open URL:', error);
            // Fallback
            window.open(url, '_blank');
        }
    };

    // Table Management Functions
    const handleTableChange = (rowIndex, colIndex, value) => {
        const newData = [...formData.tableData];
        newData[rowIndex] = [...newData[rowIndex]];
        newData[rowIndex][colIndex] = value;
        setFormData(prev => ({ ...prev, tableData: newData }));
    };

    const insertTableRow = (index) => {
        const colCount = formData.tableData[0].length;
        const newRow = Array(colCount).fill('');
        const newData = [...formData.tableData];
        newData.splice(index + 1, 0, newRow);
        setFormData(prev => ({ ...prev, tableData: newData }));
    };

    const insertTableColumn = (index) => {
        const newData = formData.tableData.map(row => {
            const newRow = [...row];
            newRow.splice(index + 1, 0, '');
            return newRow;
        });
        setFormData(prev => ({ ...prev, tableData: newData }));
    };

    const removeTableRow = (index) => {
        if (formData.tableData.length <= 1) return;
        setConfirmModal({
            show: true,
            message: 'هل أنت متأكد من حذف هذا الصف؟',
            onConfirm: () => {
                const newData = formData.tableData.filter((_, i) => i !== index);
                setFormData(prev => ({ ...prev, tableData: newData }));
                setConfirmModal({ show: false, message: '', onConfirm: null });
            }
        });
    };

    const removeTableColumn = (index) => {
        if (formData.tableData[0].length <= 1) return;
        setConfirmModal({
            show: true,
            message: 'هل أنت متأكد من حذف هذا العمود؟',
            onConfirm: () => {
                const newData = formData.tableData.map(row => row.filter((_, i) => i !== index));
                setFormData(prev => ({ ...prev, tableData: newData }));
                setConfirmModal({ show: false, message: '', onConfirm: null });
            }
        });
    };

    // Generate XML for Word Table using the existing table as a template
    // forAndroid: use fixed layout for LibreOffice compatibility
    const generateXMLTable = (data, templateXml, forAndroid = false) => {
        if (!data || data.length === 0) return '';

        // 1. Parse Template Styles
        let tblPr = '';
        let headerRowTemplate = ''; // The full TR string minus cells? No, we need properties.
        let dataRowTemplate = '';
        let headerCellTemplate = '';
        let dataCellTemplate = '';
        let totalTableWidth = 10200; // Default fallback

        if (templateXml) {
            // Extract Table Properties
            const tblPrMatch = templateXml.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/);
            if (tblPrMatch) tblPr = tblPrMatch[0];

            // Calculate Total Width from Grid
            const gridCols = templateXml.match(/<w:gridCol w:w="(\d+)"/g);
            if (gridCols) {
                totalTableWidth = gridCols.reduce((sum, col) => {
                    const match = col.match(/w:w="(\d+)"/);
                    return sum + (match ? parseInt(match[1]) : 0);
                }, 0);
            }

            // Extract Rows
            const rows = templateXml.match(/<w:tr(?:>|\s)[\s\S]*?<\/w:tr>/g);
            if (rows && rows.length > 0) {
                // Header Row (Row 0)
                headerRowTemplate = rows[0];
                // Find first cell in header row
                const firstHeaderCell = headerRowTemplate.match(/<w:tc(?:>|\s)[\s\S]*?<\/w:tc>/);
                if (firstHeaderCell) headerCellTemplate = firstHeaderCell[0];

                // Data Row (Row 1 or fallback to Row 0)
                const dataRowSource = rows.length > 1 ? rows[1] : rows[0];
                dataRowTemplate = dataRowSource;
                const firstDataCell = dataRowSource.match(/<w:tc(?:>|\s)[\s\S]*?<\/w:tc>/);
                if (firstDataCell) dataCellTemplate = firstDataCell[0];
            }
        }

        // Fallbacks if extraction failed (should not happen with valid template)
        if (!headerCellTemplate) return ''; // Fail safe

        // 2. Prepare for Generation
        const colCount = data[0].length;

        // Find the index of the "ملاحظة" column to give it special width treatment
        let noteColIndex = -1;
        data[0].forEach((cellText, index) => {
            if (cellText && cellText.includes('ملاحظة')) {
                noteColIndex = index;
            }
        });

        // Dynamic Width Strategy:
        // Calculate width for each column based on its longest content.
        // - Non-note columns: Width based on max text length in that column
        // - Note column: Huge width to take all remaining space
        // - Table layout: autofit to allow dynamic adjustment

        const HUGE_WIDTH = 5000; // For notes column (reduced from 15000 to test)
        const CHAR_WIDTH = 180;   // Approximate twips per character (Arabic is wider)
        const MIN_WIDTH = 400;    // Minimum column width (for very short content like "1")
        const PADDING = 300;      // Cell padding in twips

        // Calculate max content length for each column
        let columnWidths = data[0].map((_, colIndex) => {
            if (colIndex === noteColIndex) {
                return HUGE_WIDTH; // Notes column always gets huge width
            }

            // Find the longest text in this column
            let maxLength = 0;
            data.forEach(row => {
                const cellText = row[colIndex] || '';
                maxLength = Math.max(maxLength, cellText.length);
            });

            // Calculate width: characters * char_width + padding, but at least MIN_WIDTH
            const calculatedWidth = (maxLength * CHAR_WIDTH) + PADDING;
            return Math.max(calculatedWidth, MIN_WIDTH);
        });

        // Helper to Create a Cell from Template
        const createCell = (template, text, isNoteCol, colIndex) => {
            let cell = template;

            // Width Logic: Use calculated DXA width for this specific column
            const preferredWidth = columnWidths[colIndex];
            const widthTag = `<w:tcW w:w="${preferredWidth}" w:type="dxa"/>`;

            cell = cell.replace(/<w:tcW(?:>|\s)[^>]*?\/>/g, widthTag);

            const safeText = (text || '')
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&apos;');

            // Extract properties from template
            const pPrMatch = cell.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
            const pPr = pPrMatch ? pPrMatch[0] : '<w:pPr><w:jc w:val="center"/></w:pPr>';

            const rPrMatch = cell.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
            const rPr = rPrMatch ? rPrMatch[0] : '';

            const newParagraph = `
                <w:p>
                    ${pPr}
                    <w:r>
                        ${rPr}
                        <w:t>${safeText}</w:t>
                    </w:r>
                </w:p>
            `;

            // Re-extract tcPr from the MODIFIED cell (preserves template styling)
            const tcPrMatchNew = cell.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/);
            let tcPr = tcPrMatchNew ? tcPrMatchNew[0] : `<w:tcPr>${widthTag}<w:vAlign w:val="center"/></w:tcPr>`;

            // For PC only: Add noWrap to non-note columns
            if (!forAndroid && !isNoteCol) {
                if (!tcPr.includes('w:noWrap')) {
                    tcPr = tcPr.replace('</w:tcPr>', '<w:noWrap/></w:tcPr>');
                }
            }

            return `<w:tc>${tcPr}${newParagraph}</w:tc>`;
        };

        // 3. Generate Rows
        let rowsXML = '';

        data.forEach((row, rowIndex) => {
            const isHeader = rowIndex === 0;
            const cellTemplate = isHeader ? headerCellTemplate : dataCellTemplate;
            const rowTemplate = isHeader ? headerRowTemplate : dataRowTemplate; // For trPr

            // cellTemplate fallback
            const finalCellTemplate = cellTemplate || headerCellTemplate;

            let cellsXML = '';
            row.forEach((cellText, colIndex) => {
                const isNoteCol = (colIndex === noteColIndex);
                cellsXML += createCell(finalCellTemplate, cellText, isNoteCol, colIndex);
            });

            // Reconstruct Row using row properties from template
            const trPrMatch = rowTemplate.match(/<w:trPr>[\s\S]*?<\/w:trPr>/);
            const trPr = trPrMatch ? trPrMatch[0] : '<w:trPr><w:jc w:val="center"/></w:trPr>';

            rowsXML += `<w:tr>${trPr}${cellsXML}</w:tr>`;
        });

        // Android (LibreOffice): Use percentage width with fixed layout for reliable rendering
        // PC (MS Word): Use percentage width with autofit for dynamic behavior
        if (forAndroid) {
            // Use 100% table width (same as PC)
            if (tblPr.match(/<w:tblW(?:>|\s)[^>]*?\/>/)) {
                tblPr = tblPr.replace(/<w:tblW(?:>|\s)[^>]*?\/>/g, '<w:tblW w:w="5000" w:type="pct"/>');
            } else {
                tblPr = tblPr.replace('<w:tblPr>', '<w:tblPr><w:tblW w:w="5000" w:type="pct"/>');
            }

            // Use fixed layout for Android/LibreOffice (more predictable widths)
            tblPr = tblPr.replace(/<w:tblLayout[^>]*\/>/g, '');
            tblPr = tblPr.replace('</w:tblPr>', '<w:tblLayout w:type="fixed"/></w:tblPr>');
        } else {
            // PC: Use percentage width with autofit
            if (tblPr.match(/<w:tblW(?:>|\s)[^>]*?\/>/)) {
                tblPr = tblPr.replace(/<w:tblW(?:>|\s)[^>]*?\/>/g, '<w:tblW w:w="5000" w:type="pct"/>');
            } else {
                tblPr = tblPr.replace('<w:tblPr>', '<w:tblPr><w:tblW w:w="5000" w:type="pct"/>');
            }

            // Remove any fixed layout (ensure autofit behavior)
            tblPr = tblPr.replace(/<w:tblLayout[^>]*\/>/g, '');
            // Add autofit layout explicitly
            if (!tblPr.includes('w:tblLayout')) {
                tblPr = tblPr.replace('</w:tblPr>', '<w:tblLayout w:type="autofit"/></w:tblPr>');
            }
        }

        // 5. Final Assembly with matching grid columns
        return `
            <w:tbl>
                ${tblPr}
                <w:tblGrid>
                    ${data[0].map((_, idx) => {
            // Grid hints match dynamically calculated cell widths
            return `<w:gridCol w:w="${columnWidths[idx]}"/>`;
        }).join('')}
                </w:tblGrid>
                ${rowsXML}
            </w:tbl>
        `;
    };

    // Generate and export document
    const exportDocument = async (format) => {
        try {
            // Check platform support
            const onElectron = isElectron();
            const onAndroid = isAndroid();

            if (!onElectron && !onAndroid) {
                alert('المنصة غير مدعومة - يرجى استخدام التطبيق على الكمبيوتر أو أندرويد');
                return;
            }

            // Set loading state for all export formats
            setIsExporting(true);
            setExportFormat(format); // Track which format is being exported
            setExportProgress(10); // Initial progress
            exportCancelledRef.current = false; // Reset cancellation flag

            // Generate Word document first (for all formats)
            let templateBuffer;
            if (onElectron) {
                templateBuffer = await window.electronAPI.readTemplate();
            } else if (onAndroid) {
                // Fetch template from assets on Android
                const response = await fetch('/template.docx');
                templateBuffer = await response.arrayBuffer();
            }
            setExportProgress(30); // Template loaded

            // Check if cancelled after loading template
            if (exportCancelledRef.current) {
                throw new Error('Export cancelled');
            }

            const zip = new PizZip(templateBuffer);

            // Get all XML files from the docx
            const xmlFiles = ['word/document.xml', 'word/header1.xml', 'word/header2.xml', 'word/header3.xml'];

            // Prepare replacement data with RTL fix for Arabic/English mixed content
            const fixArabicText = (text) => {
                if (!text) return '';

                const needsFix = /[a-zA-Z0-9\/]|\\/.test(text) || /\.{2,}/.test(text);
                let processedText = text;

                if (needsFix) {
                    const RLE = '\u202B';
                    const PDF = '\u202C';
                    const LRE = '\u202A';
                    const RLM = '\u200F';

                    const placeholders = [];
                    const ltrChars = "a-zA-Z0-9<>&\\*\\%\\$#@!+=\\\\\\\\?_\\-\\^~\\|\\[\\]\\{\\}\\(\\);:\"'";
                    const regexStr = `([${ltrChars}]+(?:[\\.][${ltrChars}]+)*)`;
                    const ltrRegex = new RegExp(regexStr, 'g');

                    const protectedText = text.replace(ltrRegex, (match) => {
                        placeholders.push(LRE + match + PDF);
                        return `__PH_${placeholders.length - 1}__`;
                    });

                    let internalProcessed = protectedText.replace(/[\/\\]/g, (match) => RLM + match + RLM);
                    internalProcessed = internalProcessed.replace(/(\.+)( ?)/g, (match, dots, space) => {
                        return RLM + dots + RLM + (space || '');
                    });

                    processedText = internalProcessed.replace(/__PH_(\d+)__/g, (match, index) => placeholders[index]);

                    if (processedText.endsWith('.')) {
                        processedText = processedText.slice(0, -1) + '.' + RLM;
                    }

                    processedText = RLE + processedText + PDF;
                }

                const lines = processedText.split('\n');
                const processedLines = lines.map(line => {
                    const needsFixLine = /[a-zA-Z0-9\/]|\\/.test(line) || /\.{2,}/.test(line);
                    let pLine = line;

                    if (needsFixLine) {
                        const RLE = '\u202B';
                        const PDF = '\u202C';
                        const LRE = '\u202A';
                        const RLM = '\u200F';
                        const placeholders = [];
                        const ltrChars = "a-zA-Z0-9<>&\\*\\%\\$#@!+=\\\\\\\\?_\\-\\^~\\|\\[\\]\\{\\}\\(\\);:\"'";
                        const regexStr = `([${ltrChars}]+(?:[\\.][${ltrChars}]+)*)`;
                        const ltrRegex = new RegExp(regexStr, 'g');

                        const protectedText = line.replace(ltrRegex, (match) => {
                            placeholders.push(LRE + match + PDF);
                            return `__PH_${placeholders.length - 1}__`;
                        });

                        let internalProcessed = protectedText.replace(/[\/\\]/g, (match) => RLM + match + RLM);
                        internalProcessed = internalProcessed.replace(/(\.+)( ?)/g, (match, dots, space) => {
                            return RLM + dots + RLM + (space || '');
                        });

                        pLine = internalProcessed.replace(/__PH_(\d+)__/g, (match, index) => placeholders[index]);
                        if (pLine.endsWith('.')) {
                            pLine = pLine.slice(0, -1) + '.' + RLM;
                        }
                        pLine = RLE + pLine + PDF;
                    }

                    return pLine
                        .replace(/&/g, '&amp;')
                        .replace(/</g, '&lt;')
                        .replace(/>/g, '&gt;')
                        .replace(/"/g, '&quot;')
                        .replace(/'/g, '&apos;')
                        .replace(/\$/g, '$$$$');
                });

                return processedLines.join('</w:t><w:br/><w:t>');
            };

            const replacements = {
                'from': fixArabicText(formData.from),
                'parent_company': fixArabicText(formData.parent_company),
                'subsidiary_company': fixArabicText(formData.subsidiary_company),
                'to': fixArabicText(formData.to),
                'to_the': fixArabicText(formData.to_the),
                'greetings': fixArabicText(formData.greetings),
                'subject_name': fixArabicText(formData.subject_name),
                'subject': fixArabicText(formData.subject),
                'ending': fixArabicText(formData.ending),
                'sign': fixArabicText(formData.sign),
                'copy_to': fixArabicText(formData.copy_to),

                'للالأوللل': fixArabicText(formData.parent_company) || ' ', // Ensure textbox not empty
                'للالثانيلل': fixArabicText(formData.subsidiary_company) || ' ', // Ensure textbox not empty
                'للالثالثلل': fixArabicText(formData.from) || ' ', // Ensure textbox not empty
                'للالاخلل': fixArabicText(formData.to),
                'للالمحترملل': fixArabicText(formData.to_the),
                'للالسلام عليكم ورحمة الله وبركاتهلل': fixArabicText(formData.greetings),
                'للالموضوعلل': fixArabicText(formData.subject_name),
                'للتحية طيبة وبعدلل': fixArabicText(formData.subject),
                'للوشكرالل': fixArabicText(formData.ending),
                'للالتوقيعلل': fixArabicText(formData.sign)
            };

            const directTextKeys = Object.keys(replacements).filter(k => /[\u0600-\u06FF]/.test(k));
            directTextKeys.sort((a, b) => b.length - a.length);
            const escapedKeys = directTextKeys.map(k => k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
            const patternAll = new RegExp(`(${escapedKeys.join('|')})`, 'g');

            // Process each XML file
            xmlFiles.forEach(xmlFile => {
                try {
                    let content = zip.file(xmlFile)?.asText();
                    if (content) {
                        if (xmlFile === 'word/document.xml') {
                            const tablePattern = /<w:tbl(?:>|\s)[\s\S]*?<\/w:tbl>/g;
                            if (formData.useTable) {
                                content = content.replace(tablePattern, (match) => {
                                    if (match.includes('ملاحظة')) {
                                        return generateXMLTable(formData.tableData, match, onAndroid);
                                    }
                                    return match;
                                });
                            } else {
                                content = content.replace(tablePattern, (match) => {
                                    if (match.includes('ملاحظة')) {
                                        return '';
                                    }
                                    return match;
                                });
                            }
                        }

                        const copyToLines = formData.copy_to.split('\n').filter(line => line.trim());
                        if (copyToLines.length > 0) {
                            const paragraphPattern = /<w:p\s[^>]*>(?:[^<]|<(?!w:p\s))*?<w:numPr>(?:[^<]|<(?!w:p\s))*?<w:t[^>]*>للدائرلل<\/w:t><\/w:r><\/w:p>/;
                            const match = content.match(paragraphPattern);
                            if (match) {
                                const originalParagraph = match[0];
                                const listItems = copyToLines.map(line => {
                                    const escapedLine = line.trim()
                                        .replace(/&/g, '&amp;')
                                        .replace(/</g, '&lt;')
                                        .replace(/>/g, '&gt;')
                                        .replace(/"/g, '&quot;')
                                        .replace(/'/g, '&apos;');
                                    return originalParagraph.replace(/>للدائرلل</, '>' + escapedLine + '<');
                                }).join('');
                                content = content.replace(paragraphPattern, listItems);
                            } else {
                                const simplePattern = /(<w:t[^>]*>)للدائرلل(<\/w:t>)/g;
                                const firstLine = copyToLines[0].trim()
                                    .replace(/&/g, '&amp;')
                                    .replace(/</g, '&lt;')
                                    .replace(/>/g, '&gt;')
                                    .replace(/"/g, '&quot;')
                                    .replace(/'/g, '&apos;');
                                content = content.replace(simplePattern, `$1${firstLine}$2`);
                            }
                        } else {
                            // Remove the placeholder when copy_to is empty
                            const paragraphPattern = /<w:p\s[^>]*>(?:[^<]|<(?!w:p\s))*?<w:numPr>(?:[^<]|<(?!w:p\s))*?<w:t[^>]*>للدائرلل<\/w:t><\/w:r><\/w:p>/;
                            content = content.replace(paragraphPattern, '');
                            const simplePattern = /(<w:t[^>]*>)للدائرلل(<\/w:t>)/g;
                            content = content.replace(simplePattern, '$1$2');
                        }

                        Object.keys(replacements).forEach(key => {
                            const pattern1 = new RegExp(`(<w:t[^>]*>)«${key}»(</w:t>)`, 'g');
                            content = content.replace(pattern1, `$1${replacements[key]}$2`);
                        });

                        content = content.replace(/(<w:t[^>]*>)(.*?)(<\/w:t>)/g, (fullMatch, openTag, textContent, closeTag) => {
                            const newText = textContent.replace(patternAll, (matchedKey) => replacements[matchedKey]);
                            return `${openTag}${newText}${closeTag}`;
                        });

                        content = content.replace(/<w:t[^>]*>«<\/w:t>/g, '<w:t></w:t>');
                        content = content.replace(/<w:t[^>]*>»<\/w:t>/g, '<w:t></w:t>');

                        // Remove [[ and ]] brackets from separate XML runs (Word splits them from placeholder text)
                        content = content.replace(/<w:t[^>]*>\[\[<\/w:t>/g, '<w:t></w:t>');
                        content = content.replace(/<w:t[^>]*>\]\]<\/w:t>/g, '<w:t></w:t>');

                        zip.file(xmlFile, content);
                    }
                } catch (e) {
                    console.log(`Skipping ${xmlFile}:`, e.message);
                }
            });

            // Replace images if logo is selected
            if (formData.logoBase64) {
                try {
                    const imgBinary = atob(formData.logoBase64);
                    const imgArray = new Uint8Array(imgBinary.length);
                    for (let i = 0; i < imgBinary.length; i++) {
                        imgArray[i] = imgBinary.charCodeAt(i);
                    }
                    // Only replace image1 (Watermark) and image2 (Header Logo)
                    ['word/media/image1.png', 'word/media/image2.png'].forEach(imgName => {
                        if (zip.file(imgName)) {
                            zip.file(imgName, imgArray);
                        }
                    });
                } catch (e) {
                    console.error('Error replacing images:', e);
                }
            }

            // Remove mail merge data source connections
            try {
                let settings = zip.file('word/settings.xml')?.asText();
                if (settings) {
                    settings = settings.replace(/<w:mailMerge>[\s\S]*?<\/w:mailMerge>/g, '');
                    zip.file('word/settings.xml', settings);
                }
                if (zip.file('word/recipientData.xml')) {
                    zip.remove('word/recipientData.xml');
                }
            } catch (e) {
                console.log('Could not remove mail merge settings:', e.message);
            }

            // Remove date textbox from header3.xml if showDate is false
            if (!formData.showDate) {
                try {
                    let header3 = zip.file('word/header3.xml')?.asText();
                    if (header3) {
                        // Regex with negative lookahead to prevent matching across multiple AlternateContent blocks
                        // Only matches the specific AlternateContent containing "مستطيل 7" (date box)
                        const dateBoxPattern = /<mc:AlternateContent>(?:(?!<mc:AlternateContent>)[\s\S])*?<wp:docPr [^>]*name="مستطيل 7"[^>]*\/>[\s\S]*?<\/mc:AlternateContent>/g;
                        header3 = header3.replace(dateBoxPattern, '');
                        // Fallback: match by Anchor ID with same negative lookahead protection
                        const dateBoxPattern2 = /<mc:AlternateContent>(?:(?!<mc:AlternateContent>)[\s\S])*?wp14:anchorId="6FBCC3A1"[\s\S]*?<\/mc:AlternateContent>/g;
                        header3 = header3.replace(dateBoxPattern2, '');
                        zip.file('word/header3.xml', header3);
                    }
                } catch (e) {
                    console.log('Could not remove date textbox:', e.message);
                }
            }

            // Android-specific: Enforce A4 page size for LibreOffice compatibility
            // LibreOffice may interpret page size differently, causing tables to shrink
            // This explicitly sets page dimensions to A4 (11906 x 16838 twips)
            if (onAndroid) {
                try {
                    let docContent = zip.file('word/document.xml')?.asText();
                    if (docContent) {
                        // A4 dimensions in twips: 11906 x 16838 (portrait)
                        const a4PgSz = '<w:pgSz w:w="11906" w:h="16838"/>';

                        // Replace existing pgSz with A4 dimensions
                        if (docContent.includes('<w:pgSz')) {
                            docContent = docContent.replace(/<w:pgSz[^>]*\/>/g, a4PgSz);
                        } else {
                            // Insert pgSz into sectPr if it doesn't exist
                            docContent = docContent.replace(/<w:sectPr([^>]*)>/g, `<w:sectPr$1>${a4PgSz}`);
                        }

                        zip.file('word/document.xml', docContent);
                        console.log('Enforced A4 page size for Android/LibreOffice');
                    }
                } catch (e) {
                    console.log('Could not enforce A4 page size:', e.message);
                }
            }

            // Add user-added images to the document (each on a new page)
            // Only for Word export - PDF/Image exports handle added images separately
            if (addedImages.length > 0 && format === 'docx') {
                try {
                    let docContent = zip.file('word/document.xml')?.asText();
                    let relsContent = zip.file('word/_rels/document.xml.rels')?.asText();

                    if (docContent && relsContent) {
                        // A4 dimensions in EMUs (English Metric Units): 1 inch = 914400 EMUs
                        // A4 is 210mm x 297mm = 8.27" x 11.69"
                        // With 1 inch margins: 6.27" x 9.69" usable area
                        const maxWidthEmu = Math.floor(6.27 * 914400); // ~5,733,288 EMUs
                        const maxHeightEmu = Math.floor(9.69 * 914400); // ~8,858,496 EMUs

                        // Find the highest existing image relationship ID
                        let maxRId = 0;
                        const rIdMatches = relsContent.matchAll(/Id="rId(\d+)"/g);
                        for (const match of rIdMatches) {
                            maxRId = Math.max(maxRId, parseInt(match[1]));
                        }

                        // Find the highest existing image number in media folder
                        let maxImageNum = 0;
                        const mediaFiles = Object.keys(zip.files).filter(f => f.startsWith('word/media/image'));
                        mediaFiles.forEach(f => {
                            const match = f.match(/image(\d+)/);
                            if (match) maxImageNum = Math.max(maxImageNum, parseInt(match[1]));
                        });

                        // Process each added image (with compression to reduce memory)
                        const imageInserts = [];
                        for (let i = 0; i < addedImages.length; i++) {
                            let imgDataUrl = addedImages[i];

                            // Compress image to reduce memory usage (especially for many images)
                            try {
                                imgDataUrl = await compressImage(imgDataUrl, 0.75);
                            } catch (e) {
                                // Use original if compression fails
                            }

                            const base64Data = imgDataUrl.split(',')[1];
                            const mimeMatch = imgDataUrl.match(/data:([^;]+);/);
                            const mimeType = mimeMatch ? mimeMatch[1] : 'image/png';
                            const ext = mimeType.includes('jpeg') || mimeType.includes('jpg') ? 'jpeg' : 'png';

                            // Add image to media folder
                            const imageNum = maxImageNum + i + 1;
                            const imagePath = `word/media/addedImage${imageNum}.${ext}`;
                            const binaryString = atob(base64Data);
                            let bytes = new Uint8Array(binaryString.length);
                            for (let j = 0; j < binaryString.length; j++) {
                                bytes[j] = binaryString.charCodeAt(j);
                            }
                            zip.file(imagePath, bytes);

                            // Free memory immediately after adding to zip
                            bytes = null;

                            // Add relationship
                            const rId = maxRId + i + 1;
                            const relEntry = `<Relationship Id="rId${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/addedImage${imageNum}.${ext}"/>`;

                            // Get image dimensions
                            let imgWidth = maxWidthEmu;
                            let imgHeight = maxHeightEmu;

                            // Try to get actual dimensions from base64
                            try {
                                const img = new Image();
                                await new Promise((resolve, reject) => {
                                    img.onload = resolve;
                                    img.onerror = reject;
                                    img.src = imgDataUrl;
                                });

                                const aspectRatio = img.width / img.height;
                                const maxAspectRatio = maxWidthEmu / maxHeightEmu;

                                if (aspectRatio > maxAspectRatio) {
                                    // Image is wider - constrain by width
                                    imgWidth = maxWidthEmu;
                                    imgHeight = Math.floor(maxWidthEmu / aspectRatio);
                                } else {
                                    // Image is taller - constrain by height
                                    imgHeight = maxHeightEmu;
                                    imgWidth = Math.floor(maxHeightEmu * aspectRatio);
                                }
                            } catch (e) {
                                // Use default full-page size
                            }

                            imageInserts.push({ rId, relEntry, imgWidth, imgHeight, imageNum });

                            // Update progress for many images
                            if (addedImages.length > 5) {
                                setExportProgress(Math.floor((i + 1) / addedImages.length * 50)); // 0-50% for image processing
                            }
                        }

                        // Add all relationships
                        relsContent = relsContent.replace('</Relationships>',
                            imageInserts.map(i => i.relEntry).join('') + '</Relationships>');
                        zip.file('word/_rels/document.xml.rels', relsContent);

                        // Create image paragraphs with page breaks
                        let imageParagraphs = '';
                        for (const img of imageInserts) {
                            imageParagraphs += `
                            <w:p>
                                <w:pPr><w:pageBreakBefore/><w:jc w:val="center"/></w:pPr>
                                <w:r>
                                    <w:drawing>
                                        <wp:inline distT="0" distB="0" distL="0" distR="0">
                                            <wp:extent cx="${img.imgWidth}" cy="${img.imgHeight}"/>
                                            <wp:effectExtent l="0" t="0" r="0" b="0"/>
                                            <wp:docPr id="${img.imageNum}" name="AddedImage${img.imageNum}"/>
                                            <wp:cNvGraphicFramePr>
                                                <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                                            </wp:cNvGraphicFramePr>
                                            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                                    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                                        <pic:nvPicPr>
                                                            <pic:cNvPr id="${img.imageNum}" name="AddedImage${img.imageNum}"/>
                                                            <pic:cNvPicPr/>
                                                        </pic:nvPicPr>
                                                        <pic:blipFill>
                                                            <a:blip r:embed="rId${img.rId}"/>
                                                            <a:stretch><a:fillRect/></a:stretch>
                                                        </pic:blipFill>
                                                        <pic:spPr>
                                                            <a:xfrm>
                                                                <a:off x="0" y="0"/>
                                                                <a:ext cx="${img.imgWidth}" cy="${img.imgHeight}"/>
                                                            </a:xfrm>
                                                            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                                                        </pic:spPr>
                                                    </pic:pic>
                                                </a:graphicData>
                                            </a:graphic>
                                        </wp:inline>
                                    </w:drawing>
                                </w:r>
                            </w:p>`;
                        }

                        // Insert before the final sectPr (section properties)
                        docContent = docContent.replace(/(<w:sectPr)/, imageParagraphs + '$1');
                        zip.file('word/document.xml', docContent);
                    }
                } catch (e) {
                    console.error('Failed to add images to DOCX:', e);
                    // Continue without added images
                }
            }

            const output = zip.generate({
                type: 'arraybuffer',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });

            const docxBuffer = Array.from(new Uint8Array(output));
            setExportProgress(50); // Document generated

            // Check if cancelled before proceeding with save/conversion
            if (exportCancelledRef.current) {
                throw new Error('Export cancelled');
            }

            setExportProgress(70); // Starting format-specific export
            // Handle different export formats
            if (format === 'docx') {
                // Save as Word document
                if (onAndroid) {
                    // Check cancellation before showing share dialog
                    if (exportCancelledRef.current) {
                        throw new Error('Export cancelled');
                    }
                    // For Android, use Filesystem + Share plugin
                    const filename = `${sanitizeFileName(formData.subject_name)}.docx`;
                    const saveResult = await saveFileAndroid(
                        docxBuffer,
                        filename,
                        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    );
                    if (saveResult.success) {
                        if (!saveResult.cancelled) {
                            showToast('تم حفظ المستند بنجاح');
                        }
                        // If cancelled, just close silently
                    } else {
                        throw new Error(saveResult.error || 'فشل حفظ المستند');
                    }
                } else {
                    const result = await window.electronAPI.saveDocument({
                        buffer: docxBuffer,
                        defaultName: `${sanitizeFileName(formData.subject_name)}.docx`,
                        filters: [{ name: 'Word Document', extensions: ['docx'] }]
                    });

                    if (result.success) {
                        showToast('تم حفظ المستند بنجاح');
                    }
                }
            } else if (format === 'pdf') {
                // Convert to PDF

                if (onElectron) {
                    // Use MS Word on Electron

                    if (addedImages.length > 0) {
                        // Need to merge document PDF with added images into single PDF
                        // Strategy: Convert document to images first, then create PDF with jsPDF
                        const conversionResult = await window.electronAPI.convertDocxToImageWord(
                            docxBuffer,
                            false, // Don't combine - we need individual pages
                            formData.subject_name || 'document'
                        );

                        if (conversionResult.success) {
                            // Create jsPDF with all images
                            const mergedPdf = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4' });
                            const pdfWidth = mergedPdf.internal.pageSize.getWidth();
                            const pdfHeight = mergedPdf.internal.pageSize.getHeight();

                            // Helper to add image to PDF page
                            const addImageToPage = async (imgData, isFirst) => {
                                if (!isFirst) mergedPdf.addPage();

                                let finalImgData = imgData;
                                let imgFormat = 'JPEG';

                                try {
                                    finalImgData = await compressImage(imgData, 0.8);
                                } catch (e) {
                                    // Detect format from original
                                    if (imgData.includes('data:image/png')) imgFormat = 'PNG';
                                    else if (imgData.includes('data:image/webp')) imgFormat = 'WEBP';
                                    finalImgData = imgData;
                                }

                                try {
                                    const imgProps = mergedPdf.getImageProperties(finalImgData);
                                    const imgRatio = imgProps.width / imgProps.height;
                                    const pageRatio = pdfWidth / pdfHeight;

                                    let w, h;
                                    if (imgRatio > pageRatio) {
                                        w = pdfWidth;
                                        h = w / imgRatio;
                                    } else {
                                        h = pdfHeight;
                                        w = h * imgRatio;
                                    }

                                    const x = (pdfWidth - w) / 2;
                                    const y = (pdfHeight - h) / 2;

                                    mergedPdf.addImage(finalImgData, imgFormat, x, y, w, h);
                                } catch (imgError) {
                                    console.error('Failed to add image to PDF:', imgError);
                                }
                            };

                            // Add document pages
                            for (let i = 0; i < conversionResult.images.length; i++) {
                                const pageBuffer = conversionResult.images[i].buffer;
                                const bytes = new Uint8Array(pageBuffer);
                                const blob = new Blob([bytes], { type: 'image/png' });
                                const dataUrl = await new Promise(resolve => {
                                    const reader = new FileReader();
                                    reader.onload = () => resolve(reader.result);
                                    reader.readAsDataURL(blob);
                                });
                                await addImageToPage(dataUrl, i === 0);
                            }

                            // Add added images
                            for (let i = 0; i < addedImages.length; i++) {
                                await addImageToPage(addedImages[i], false);
                            }

                            // Save merged PDF
                            const pdfBuffer = mergedPdf.output('arraybuffer');
                            const result = await window.electronAPI.saveDocument({
                                buffer: Array.from(new Uint8Array(pdfBuffer)),
                                defaultName: `${sanitizeFileName(formData.subject_name)}.pdf`,
                                filters: [{ name: 'PDF Document', extensions: ['pdf'] }]
                            });

                            if (result.success) {
                                showToast('تم حفظ ملف PDF بنجاح');
                            }
                        } else {
                            throw new Error('فشل تحويل المستند إلى PDF');
                        }
                    } else {
                        // No added images - use direct PDF conversion
                        const conversionResult = await window.electronAPI.convertDocxToPdfWord(docxBuffer);

                        if (conversionResult.success) {
                            const result = await window.electronAPI.saveDocument({
                                buffer: conversionResult.buffer,
                                defaultName: `${sanitizeFileName(formData.subject_name)}.pdf`,
                                filters: [{ name: 'PDF Document', extensions: ['pdf'] }]
                            });

                            if (result.success) {
                                showToast('تم حفظ ملف PDF بنجاح');
                            }
                        } else {
                            throw new Error('فشل تحويل المستند إلى PDF');
                        }
                    }
                } else if (onAndroid) {
                    // Use LibreOfficeKit on Android

                    if (addedImages.length > 0) {
                        // Need to merge document PDF with added images into single PDF
                        // Strategy: Convert document to images first, then create PDF with jsPDF
                        const conversionResult = await convertToImage(docxBuffer, false, formData.subject_name);

                        if (conversionResult.success) {
                            // Check cancellation
                            if (exportCancelledRef.current) {
                                throw new Error('Export cancelled');
                            }

                            // Create jsPDF with all images
                            const mergedPdf = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4' });
                            const pdfWidth = mergedPdf.internal.pageSize.getWidth();
                            const pdfHeight = mergedPdf.internal.pageSize.getHeight();

                            // Helper to add image to PDF page
                            const addImageToPage = async (imgData, isFirst) => {
                                if (!isFirst) mergedPdf.addPage();

                                let finalImgData = imgData;
                                let imgFormat = 'JPEG';

                                try {
                                    finalImgData = await compressImage(imgData, 0.8);
                                } catch (e) {
                                    // Detect format from original
                                    if (imgData.includes('data:image/png')) imgFormat = 'PNG';
                                    else if (imgData.includes('data:image/webp')) imgFormat = 'WEBP';
                                    finalImgData = imgData;
                                }

                                try {
                                    const imgProps = mergedPdf.getImageProperties(finalImgData);
                                    const imgRatio = imgProps.width / imgProps.height;
                                    const pageRatio = pdfWidth / pdfHeight;

                                    let w, h;
                                    if (imgRatio > pageRatio) {
                                        w = pdfWidth;
                                        h = w / imgRatio;
                                    } else {
                                        h = pdfHeight;
                                        w = h * imgRatio;
                                    }

                                    const x = (pdfWidth - w) / 2;
                                    const y = (pdfHeight - h) / 2;

                                    mergedPdf.addImage(finalImgData, imgFormat, x, y, w, h);
                                } catch (imgError) {
                                    console.error('Failed to add image to PDF:', imgError);
                                }
                            };

                            // Add document pages
                            for (let i = 0; i < conversionResult.images.length; i++) {
                                const pageBuffer = conversionResult.images[i];
                                const bytes = new Uint8Array(pageBuffer);
                                const blob = new Blob([bytes], { type: 'image/png' });
                                const dataUrl = await new Promise(resolve => {
                                    const reader = new FileReader();
                                    reader.onload = () => resolve(reader.result);
                                    reader.readAsDataURL(blob);
                                });
                                await addImageToPage(dataUrl, i === 0);
                            }

                            // Add added images
                            for (let i = 0; i < addedImages.length; i++) {
                                await addImageToPage(addedImages[i], false);
                            }

                            // Save merged PDF
                            const pdfBuffer = mergedPdf.output('arraybuffer');
                            const filename = `${sanitizeFileName(formData.subject_name)}.pdf`;
                            const saveResult = await saveFileAndroid(Array.from(new Uint8Array(pdfBuffer)), filename, 'application/pdf');

                            if (saveResult.success) {
                                if (!saveResult.cancelled) {
                                    showToast('تم حفظ ملف PDF بنجاح');
                                }
                            } else {
                                throw new Error(saveResult.error || 'فشل حفظ ملف PDF');
                            }
                        } else {
                            throw new Error(conversionResult.error || 'فشل تحويل المستند إلى PDF');
                        }
                    } else {
                        // No added images - use direct PDF conversion
                        const conversionResult = await convertToPdf(docxBuffer);

                        if (conversionResult.success) {
                            if (exportCancelledRef.current) {
                                throw new Error('Export cancelled');
                            }
                            setExportProgress(85); // Conversion complete

                            const filename = `${sanitizeFileName(formData.subject_name)}.pdf`;
                            setExportProgress(95); // Starting save
                            const saveResult = await saveFileAndroid(conversionResult.buffer, filename, 'application/pdf');
                            if (saveResult.success) {
                                if (!saveResult.cancelled) {
                                    showToast('تم حفظ ملف PDF بنجاح');
                                }
                            } else {
                                throw new Error(saveResult.error || 'فشل حفظ ملف PDF');
                            }
                        } else {
                            throw new Error(conversionResult.error || 'فشل تحويل المستند إلى PDF');
                        }
                    }
                }
            } else if (format === 'png') {
                // Convert to Image
                // Convert to Image

                if (onAndroid) {
                    // Use LibreOfficeKit for image conversion on Android
                    const conversionResult = await convertToImage(docxBuffer, combineImagePages, formData.subject_name);

                    if (conversionResult.success) {
                        // Check cancellation before showing share dialog
                        if (exportCancelledRef.current) {
                            throw new Error('Export cancelled');
                        }
                        setExportProgress(85); // Conversion complete

                        if (combineImagePages) {
                            let finalBuffer = conversionResult.buffer;

                            // If there are added images, merge them with the document image
                            if (addedImages.length > 0) {
                                // Helper to load image from data URL or buffer
                                const loadImg = (src) => new Promise((resolve, reject) => {
                                    const img = new Image();
                                    img.onload = () => resolve(img);
                                    img.onerror = reject;
                                    img.src = src;
                                });

                                // Convert buffer to data URL
                                const docBytes = new Uint8Array(conversionResult.buffer);
                                const docBlob = new Blob([docBytes], { type: 'image/png' });
                                const docDataUrl = await new Promise(resolve => {
                                    const reader = new FileReader();
                                    reader.onload = () => resolve(reader.result);
                                    reader.readAsDataURL(docBlob);
                                });

                                // Load all images (document + added)
                                const allImages = [docDataUrl, ...addedImages];
                                const loadedImages = await Promise.all(allImages.map(loadImg));

                                // Calculate total dimensions
                                let totalHeight = 0;
                                let maxWidth = 0;
                                for (const img of loadedImages) {
                                    totalHeight += img.height;
                                    maxWidth = Math.max(maxWidth, img.width);
                                }

                                // Create canvas and draw all images
                                const canvas = document.createElement('canvas');
                                canvas.width = maxWidth;
                                canvas.height = totalHeight;
                                const ctx = canvas.getContext('2d');
                                ctx.fillStyle = '#FFFFFF';
                                ctx.fillRect(0, 0, maxWidth, totalHeight);

                                let yOffset = 0;
                                for (const img of loadedImages) {
                                    const xOffset = (maxWidth - img.width) / 2; // Center horizontally
                                    ctx.drawImage(img, xOffset, yOffset);
                                    yOffset += img.height;
                                }

                                // Convert canvas to buffer
                                const dataUrl = canvas.toDataURL('image/png');
                                const base64 = dataUrl.split(',')[1];
                                const binary = atob(base64);
                                const bytes = new Uint8Array(binary.length);
                                for (let i = 0; i < binary.length; i++) {
                                    bytes[i] = binary.charCodeAt(i);
                                }
                                finalBuffer = Array.from(bytes);
                            }

                            // Single merged image
                            const filename = `${sanitizeFileName(formData.subject_name)}.png`;
                            setExportProgress(95); // Starting save
                            const saveResult = await saveFileAndroid(finalBuffer, filename, 'image/png');
                            if (saveResult.success) {
                                if (!saveResult.cancelled) {
                                    showToast('تم حفظ الصورة بنجاح');
                                }
                            } else {
                                throw new Error(saveResult.error || 'فشل حفظ الصورة');
                            }
                        } else {
                            // Multiple separate images - share all with one dialog
                            const images = conversionResult.images;
                            const baseName = sanitizeFileName(formData.subject_name);

                            // Prepare files array from document pages
                            const files = images.map((buffer, i) => ({
                                buffer: buffer,
                                filename: `${baseName}_page${i + 1}.png`
                            }));

                            // Append added images (convert base64 to buffer)
                            if (addedImages.length > 0) {
                                for (let i = 0; i < addedImages.length; i++) {
                                    const base64 = addedImages[i].split(',')[1]; // Remove data:image/...;base64, prefix
                                    const binary = atob(base64);
                                    const bytes = new Uint8Array(binary.length);
                                    for (let j = 0; j < binary.length; j++) {
                                        bytes[j] = binary.charCodeAt(j);
                                    }
                                    files.push({
                                        buffer: bytes,
                                        filename: `${baseName}_added${i + 1}.png`
                                    });
                                }
                            }

                            // Check cancellation before showing share dialog
                            if (exportCancelledRef.current) {
                                throw new Error('Export cancelled');
                            }
                            const saveResult = await saveMultipleFilesAndroid(files, 'حفظ أو مشاركة الصور');
                            if (saveResult.success) {
                                if (!saveResult.cancelled) {
                                    showToast(`تم مشاركة ${saveResult.count} صورة بنجاح`);
                                }
                            } else {
                                throw new Error(saveResult.error || 'فشل حفظ الصورة');
                            }
                        }
                    } else {
                        throw new Error(conversionResult.error || 'فشل تحويل المستند إلى صورة');
                    }
                } else if (onElectron) {
                    // Use MS Word + pdf-poppler on Electron

                    const conversionResult = await window.electronAPI.convertDocxToImageWord(
                        docxBuffer,
                        combineImagePages,
                        formData.subject_name || 'document'
                    );

                    if (conversionResult.success) {
                        if (combineImagePages) {
                            let finalBuffer = conversionResult.buffer;

                            // If there are added images, merge them with the document image
                            if (addedImages.length > 0) {
                                // Helper to load image from data URL or buffer
                                const loadImg = (src) => new Promise((resolve, reject) => {
                                    const img = new Image();
                                    img.onload = () => resolve(img);
                                    img.onerror = reject;
                                    img.src = src;
                                });

                                // Convert buffer to data URL
                                const docBytes = new Uint8Array(conversionResult.buffer);
                                const docBlob = new Blob([docBytes], { type: 'image/png' });
                                const docDataUrl = await new Promise(resolve => {
                                    const reader = new FileReader();
                                    reader.onload = () => resolve(reader.result);
                                    reader.readAsDataURL(docBlob);
                                });

                                // Load all images (document + added)
                                const allImages = [docDataUrl, ...addedImages];
                                const loadedImages = await Promise.all(allImages.map(loadImg));

                                // Calculate total dimensions
                                let totalHeight = 0;
                                let maxWidth = 0;
                                for (const img of loadedImages) {
                                    totalHeight += img.height;
                                    maxWidth = Math.max(maxWidth, img.width);
                                }

                                // Create canvas and draw all images
                                const canvas = document.createElement('canvas');
                                canvas.width = maxWidth;
                                canvas.height = totalHeight;
                                const ctx = canvas.getContext('2d');
                                ctx.fillStyle = '#FFFFFF';
                                ctx.fillRect(0, 0, maxWidth, totalHeight);

                                let yOffset = 0;
                                for (const img of loadedImages) {
                                    const xOffset = (maxWidth - img.width) / 2; // Center horizontally
                                    ctx.drawImage(img, xOffset, yOffset);
                                    yOffset += img.height;
                                }

                                // Convert canvas to buffer
                                const dataUrl = canvas.toDataURL('image/png');
                                const base64 = dataUrl.split(',')[1];
                                const binary = atob(base64);
                                const bytes = new Uint8Array(binary.length);
                                for (let i = 0; i < binary.length; i++) {
                                    bytes[i] = binary.charCodeAt(i);
                                }
                                finalBuffer = Array.from(bytes);
                            }

                            // Save the combined image
                            const result = await window.electronAPI.saveDocument({
                                buffer: finalBuffer,
                                defaultName: `${sanitizeFileName(formData.subject_name)}.png`,
                                filters: [{ name: 'PNG Image', extensions: ['png'] }]
                            });

                            if (result.success) {
                                showToast('تم حفظ الصورة بنجاح');
                            }
                        } else {
                            // Multiple separate images - save all with one dialog
                            const images = [...conversionResult.images];

                            // Append added images (convert base64 to buffer array)
                            if (addedImages.length > 0) {
                                for (let i = 0; i < addedImages.length; i++) {
                                    const base64 = addedImages[i].split(',')[1]; // Remove prefix
                                    const binary = atob(base64);
                                    const bytes = new Uint8Array(binary.length);
                                    for (let j = 0; j < binary.length; j++) {
                                        bytes[j] = binary.charCodeAt(j);
                                    }
                                    images.push({ buffer: Array.from(bytes) });
                                }
                            }

                            const result = await window.electronAPI.saveMultipleImages(
                                images,
                                formData.subject_name || 'document'
                            );

                            if (result.success) {
                                showToast(`تم حفظ ${result.count} صورة بنجاح`);
                            }
                        }
                    } else {
                        throw new Error('فشل تحويل المستند إلى صورة');
                    }
                }
            }
        } catch (error) {
            console.error('Export error:', error);
            // Don't show error for cancelled operations
            if (!error.message?.includes('cancelled')) {
                const errorMsg = error.message || error.toString() || 'خطأ غير معروف';
                alert('حدث خطأ أثناء التصدير: ' + errorMsg);
            }
        } finally {
            // Reset loading state (keep addedImages for potential re-export)
            setIsExporting(false);
            setExportProgress(0); // Reset progress
        }
    };



    // External: Export to PDF
    const handleExternalExportPdf = async () => {
        if (previewState.images.length === 0) return;

        setIsExporting(true);
        setExportFormat('pdf'); // Enable cancel button
        setExportProgress(10); // Initial progress
        exportCancelledRef.current = false;
        try {
            const doc = new jsPDF({
                orientation: 'p',
                unit: 'mm',
                format: 'a4'
            }); // Default A4 portrait

            const pdfWidth = doc.internal.pageSize.getWidth();
            const pdfHeight = doc.internal.pageSize.getHeight();
            const totalImages = previewState.images.length;

            for (let i = 0; i < previewState.images.length; i++) {
                // Check cancellation
                if (exportCancelledRef.current) throw new Error('CANCELLED');

                // Update progress based on image index (10-90%)
                setExportProgress(10 + Math.round((i / totalImages) * 80));

                if (i > 0) doc.addPage();

                const imgData = previewState.images[i];

                // Compress image to reduce file size
                let finalImgData = imgData;
                let format = 'JPEG';

                try {
                    // Always compress to JPEG 0.8 to reduce size
                    finalImgData = await compressImage(imgData, 0.8);
                } catch (e) {
                    console.error('Image compression failed, using original', e);
                    // Detect format from original data URL
                    if (imgData.includes('data:image/png')) {
                        format = 'PNG';
                    } else if (imgData.includes('data:image/webp')) {
                        format = 'WEBP';
                    } else if (imgData.includes('data:image/gif')) {
                        format = 'GIF';
                    }
                    finalImgData = imgData;
                }

                try {
                    // Get image dimensions to fit to page
                    const imgProps = doc.getImageProperties(finalImgData);

                    // Calculate aspect ratio
                    const imgRatio = imgProps.width / imgProps.height;
                    const pdfRatio = pdfWidth / pdfHeight;

                    let w, h;
                    if (imgRatio > pdfRatio) {
                        w = pdfWidth;
                        h = w / imgRatio;
                    } else {
                        h = pdfHeight;
                        w = h * imgRatio;
                    }

                    // Center image
                    const x = (pdfWidth - w) / 2;
                    const y = (pdfHeight - h) / 2;

                    doc.addImage(finalImgData, format, x, y, w, h);
                } catch (imgError) {
                    console.error('Failed to add image to PDF:', imgError);
                    // Skip this image but continue with others
                }

                // Yield to allow UI update
                await new Promise(resolve => setTimeout(resolve, 0));
            }

            if (exportCancelledRef.current) throw new Error('CANCELLED');

            const pdfArrayBuffer = doc.output('arraybuffer');

            // Platform-specific saving
            if (isElectron() && window.electronAPI) {
                // PC: Use Electron save dialog
                const filename = `exported_${Date.now()}.pdf`;
                const saveResult = await window.electronAPI.saveDocument({
                    buffer: Array.from(new Uint8Array(pdfArrayBuffer)),
                    defaultName: filename,
                    filters: [{ name: 'PDF Document', extensions: ['pdf'] }]
                });

                if (saveResult.success) {
                    showToast('تم تصدير PDF بنجاح');
                } else if (!saveResult.cancelled) {
                    // User cancelled the dialog - do nothing
                }
            } else if (isAndroid()) {
                // Android: Save to cache first then share
                const filename = `shared_export_${Date.now()}.pdf`;
                const saveResult = await saveToCacheAndroid(pdfArrayBuffer, filename);

                if (!saveResult.success) throw new Error(saveResult.error || 'Failed to save');

                // Check cancellation AGAIN before showing share dialog
                if (exportCancelledRef.current) throw new Error('CANCELLED');

                // Share logic
                const shareResult = await shareFileAndroid(saveResult.uri, filename, 'application/pdf');

                if (shareResult.success) {
                    if (!shareResult.cancelled) showToast('تم تصدير PDF بنجاح');
                } else {
                    throw new Error(shareResult.error || 'Failed');
                }
            } else {
                // Fallback: trigger browser download
                const blob = new Blob([pdfArrayBuffer], { type: 'application/pdf' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `exported_${Date.now()}.pdf`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                showToast('تم تصدير PDF بنجاح');
            }

        } catch (e) {
            if (e.message === 'CANCELLED') {
                showToast('تم إلغاء التصدير', 'info');
            } else {
                console.error(e);
                showToast('خطأ في التصدير: ' + e.message, 'error');
            }
        } finally {
            setIsExporting(false);
            setExportProgress(0); // Reset progress
        }
    };

    // Handle document preview - generates images and shows in overlay
    const handlePreview = async () => {
        try {
            const onElectron = isElectron();
            const onAndroid = isAndroid();

            if (!onElectron && !onAndroid) {
                showToast('المنصة غير مدعومة', 'error');
                return;
            }

            // Set loading state
            setPreviewState(prev => ({ ...prev, show: true, isLoading: true, images: [], currentPage: 0, isExternal: false }));

            // Generate Word document (same logic as exportDocument)
            let templateBuffer;
            if (onElectron) {
                templateBuffer = await window.electronAPI.readTemplate();
            } else if (onAndroid) {
                const response = await fetch('/template.docx');
                templateBuffer = await response.arrayBuffer();
            }

            const zip = new PizZip(templateBuffer);
            const xmlFiles = ['word/document.xml', 'word/header1.xml', 'word/header2.xml', 'word/header3.xml'];

            // Prepare replacement data with RTL fix
            const fixArabicText = (text) => {
                if (!text) return '';
                const needsFix = /[a-zA-Z0-9\/]|\\/.test(text) || /\.{2,}/.test(text);
                let processedText = text;
                if (needsFix) {
                    const RLE = '\u202B'; const PDF = '\u202C'; const LRE = '\u202A'; const RLM = '\u200F';
                    const placeholders = [];
                    const ltrChars = "a-zA-Z0-9<>&\\*\\%\\$#@!+=\\\\\\\\?_\\-\\^~\\|\\[\\]\\{\\}\\(\\);:\"'";
                    const ltrRegex = new RegExp(`([${ltrChars}]+(?:[\\.][${ltrChars}]+)*)`, 'g');
                    const protectedText = text.replace(ltrRegex, (match) => { placeholders.push(LRE + match + PDF); return `__PH_${placeholders.length - 1}__`; });
                    let internalProcessed = protectedText.replace(/[\/\\]/g, (match) => RLM + match + RLM);
                    internalProcessed = internalProcessed.replace(/(\.+)( ?)/g, (m, dots, space) => RLM + dots + RLM + (space || ''));
                    processedText = internalProcessed.replace(/__PH_(\d+)__/g, (m, index) => placeholders[index]);
                    if (processedText.endsWith('.')) processedText = processedText.slice(0, -1) + '.' + RLM;
                    processedText = RLE + processedText + PDF;
                }
                const lines = processedText.split('\n');
                const processedLines = lines.map(line => {
                    return line.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;').replace(/\$/g, '$$$$');
                });
                return processedLines.join('</w:t><w:br/><w:t>');
            };

            const replacements = {
                'from': fixArabicText(formData.from), 'parent_company': fixArabicText(formData.parent_company),
                'subsidiary_company': fixArabicText(formData.subsidiary_company), 'to': fixArabicText(formData.to),
                'to_the': fixArabicText(formData.to_the), 'greetings': fixArabicText(formData.greetings),
                'subject_name': fixArabicText(formData.subject_name), 'subject': fixArabicText(formData.subject),
                'ending': fixArabicText(formData.ending), 'sign': fixArabicText(formData.sign), 'copy_to': fixArabicText(formData.copy_to),
                'للالأوللل': fixArabicText(formData.parent_company) || ' ',
                'للالثانيلل': fixArabicText(formData.subsidiary_company) || ' ',
                'للالثالثلل': fixArabicText(formData.from) || ' ',
                'للالاخلل': fixArabicText(formData.to), 'للالمحترملل': fixArabicText(formData.to_the),
                'للالسلام عليكم ورحمة الله وبركاتهلل': fixArabicText(formData.greetings),
                'للالموضوعلل': fixArabicText(formData.subject_name),
                'للتحية طيبة وبعدلل': fixArabicText(formData.subject), 'للوشكرالل': fixArabicText(formData.ending), 'للالتوقيعلل': fixArabicText(formData.sign)
            };

            const directTextKeys = Object.keys(replacements).filter(k => /[\u0600-\u06FF]/.test(k));
            directTextKeys.sort((a, b) => b.length - a.length);
            const escapedKeys = directTextKeys.map(k => k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
            const patternAll = new RegExp(`(${escapedKeys.join('|')})`, 'g');

            // Process XML files (simplified version of export logic)
            xmlFiles.forEach(xmlFile => {
                try {
                    let content = zip.file(xmlFile)?.asText();
                    if (content) {
                        if (xmlFile === 'word/document.xml') {
                            const tablePattern = /<w:tbl(?:>|\s)[\s\S]*?<\/w:tbl>/g;
                            if (formData.useTable) {
                                content = content.replace(tablePattern, (match) => match.includes('ملاحظة') ? generateXMLTable(formData.tableData, match, onAndroid) : match);
                            } else {
                                content = content.replace(tablePattern, (match) => match.includes('ملاحظة') ? '' : match);
                            }
                        }

                        // Handle copy_to (نسخة الى) field - same as exportDocument
                        const copyToLines = formData.copy_to.split('\n').filter(line => line.trim());
                        if (copyToLines.length > 0) {
                            const paragraphPattern = /<w:p\s[^>]*>(?:[^<]|<(?!w:p\s))*?<w:numPr>(?:[^<]|<(?!w:p\s))*?<w:t[^>]*>للدائرلل<\/w:t><\/w:r><\/w:p>/;
                            const match = content.match(paragraphPattern);
                            if (match) {
                                const originalParagraph = match[0];
                                const listItems = copyToLines.map(line => {
                                    const escapedLine = line.trim()
                                        .replace(/&/g, '&amp;')
                                        .replace(/</g, '&lt;')
                                        .replace(/>/g, '&gt;')
                                        .replace(/"/g, '&quot;')
                                        .replace(/'/g, '&apos;');
                                    return originalParagraph.replace(/>للدائرلل</, '>' + escapedLine + '<');
                                }).join('');
                                content = content.replace(paragraphPattern, listItems);
                            } else {
                                const simplePattern = /(<w:t[^>]*>)للدائرلل(<\/w:t>)/g;
                                const firstLine = copyToLines[0].trim()
                                    .replace(/&/g, '&amp;')
                                    .replace(/</g, '&lt;')
                                    .replace(/>/g, '&gt;')
                                    .replace(/"/g, '&quot;')
                                    .replace(/'/g, '&apos;');
                                content = content.replace(simplePattern, `$1${firstLine}$2`);
                            }
                        } else {
                            // Remove the placeholder when copy_to is empty
                            const paragraphPattern = /<w:p\s[^>]*>(?:[^<]|<(?!w:p\s))*?<w:numPr>(?:[^<]|<(?!w:p\s))*?<w:t[^>]*>للدائرلل<\/w:t><\/w:r><\/w:p>/;
                            content = content.replace(paragraphPattern, '');
                            const simplePattern = /(<w:t[^>]*>)للدائرلل(<\/w:t>)/g;
                            content = content.replace(simplePattern, '$1$2');
                        }

                        Object.keys(replacements).forEach(key => {
                            const pattern1 = new RegExp(`(<w:t[^>]*>)«${key}»(</w:t>)`, 'g');
                            content = content.replace(pattern1, `$1${replacements[key]}$2`);
                        });
                        content = content.replace(/(<w:t[^>]*>)(.*?)(<\/w:t>)/g, (fullMatch, openTag, textContent, closeTag) => {
                            const newText = textContent.replace(patternAll, (matchedKey) => replacements[matchedKey]);
                            return `${openTag}${newText}${closeTag}`;
                        });
                        content = content.replace(/<w:t[^>]*>«<\/w:t>/g, '<w:t></w:t>').replace(/<w:t[^>]*>»<\/w:t>/g, '<w:t></w:t>');
                        // Remove [[ and ]] brackets from separate XML runs (Word splits them from placeholder text)
                        content = content.replace(/<w:t[^>]*>\[\[<\/w:t>/g, '<w:t></w:t>');
                        content = content.replace(/<w:t[^>]*>\]\]<\/w:t>/g, '<w:t></w:t>');
                        zip.file(xmlFile, content);
                    }
                } catch (e) { console.log(`Skipping ${xmlFile}:`, e.message); }
            });

            // Replace logo if selected
            if (formData.logoBase64) {
                try {
                    const imgBinary = atob(formData.logoBase64);
                    const imgArray = new Uint8Array(imgBinary.length);
                    for (let i = 0; i < imgBinary.length; i++) imgArray[i] = imgBinary.charCodeAt(i);
                    ['word/media/image1.png', 'word/media/image2.png'].forEach(imgName => { if (zip.file(imgName)) zip.file(imgName, imgArray); });
                } catch (e) { console.error('Error replacing images:', e); }
            }

            // Remove mail merge and date box if needed
            try {
                let settings = zip.file('word/settings.xml')?.asText();
                if (settings) { settings = settings.replace(/<w:mailMerge>[\s\S]*?<\/w:mailMerge>/g, ''); zip.file('word/settings.xml', settings); }
            } catch (e) { }
            if (!formData.showDate) {
                try {
                    let header3 = zip.file('word/header3.xml')?.asText();
                    if (header3) {
                        header3 = header3.replace(/<mc:AlternateContent>(?:(?!<mc:AlternateContent>)[\s\S])*?<wp:docPr [^>]*name="مستطيل 7"[^>]*\/>[\s\S]*?<\/mc:AlternateContent>/g, '');
                        zip.file('word/header3.xml', header3);
                    }
                } catch (e) { }
            }

            const output = zip.generate({ type: 'arraybuffer', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
            const docxBuffer = Array.from(new Uint8Array(output));

            // Convert to images
            // Always separate pages for preview navigation on both platforms
            const conversionResult = await convertToImage(docxBuffer, false, formData.subject_name);

            if (conversionResult.success) {
                // Convert buffers to data URLs
                const imageDataUrls = [];
                const imagesToProcess = conversionResult.images || [conversionResult.buffer];

                for (const imageData of imagesToProcess) {
                    // Handle both Electron format (object with .buffer) and Android format (raw array)
                    const rawBuffer = imageData.buffer || imageData;
                    const bytes = new Uint8Array(rawBuffer);
                    let binary = '';
                    for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
                    const base64 = btoa(binary);
                    imageDataUrls.push(`data:image/png;base64,${base64}`);
                }

                setPreviewState(prev => ({ ...prev, images: imageDataUrls, isLoading: false }));
            } else {
                throw new Error(conversionResult.error || 'فشل في إنشاء المعاينة');
            }
        } catch (error) {
            console.error('Preview error:', error);
            setPreviewState(prev => ({ ...prev, show: false, isLoading: false }));
            showToast('حدث خطأ أثناء إنشاء المعاينة: ' + error.message, 'error');
        }
    };

    // Preview navigation helpers
    const previewNextPage = () => {
        setPreviewState(prev => ({
            ...prev,
            currentPage: Math.min(prev.currentPage + 1, prev.images.length - 1),
            zoom: 1.0,
            panX: 0,
            panY: 0
        }));
    };

    const previewPrevPage = () => {
        setPreviewState(prev => ({
            ...prev,
            currentPage: Math.max(prev.currentPage - 1, 0),
            zoom: 1.0,
            panX: 0,
            panY: 0
        }));
    };

    const previewZoomIn = () => {
        setPreviewState(prev => ({
            ...prev,
            zoom: Math.min(prev.zoom + 0.25, 3.0)
        }));
    };

    const previewZoomOut = () => {
        setPreviewState(prev => ({
            ...prev,
            zoom: Math.max(prev.zoom - 0.25, 0.5)
        }));
    };

    const closePreview = () => {
        setPreviewState({ show: false, images: [], currentPage: 0, zoom: 1.0, panX: 0, panY: 0, isLoading: false });
        setAddedImages([]); // Clear added images when closing preview
    };

    // Mouse wheel zoom handler (Zoom to Mouse)
    const handlePreviewWheel = (e) => {
        e.preventDefault();
        const rect = e.currentTarget.getBoundingClientRect();
        const cx = rect.left + rect.width / 2;
        const cy = rect.top + rect.height / 2;

        // Mouse position relative to center
        const mouseX = e.clientX - cx;
        const mouseY = e.clientY - cy;

        const delta = e.deltaY > 0 ? -0.1 : 0.1;

        setPreviewState(prev => {
            const oldZoom = prev.zoom;
            const newZoom = Math.min(Math.max(oldZoom + delta, 0.5), 3.0);

            if (newZoom === oldZoom) return prev;

            const scaleRatio = newZoom / oldZoom;
            const newPanX = mouseX - (mouseX - prev.panX) * scaleRatio;
            const newPanY = mouseY - (mouseY - prev.panY) * scaleRatio;

            return {
                ...prev,
                zoom: newZoom,
                panX: newPanX,
                panY: newPanY
            };
        });
    };

    // Pinch-to-zoom and pan state ref
    const pinchRef = useRef({ initialDistance: 0, initialZoom: 1, initialPanX: 0, initialPanY: 0, startCenterX: 0, startCenterY: 0 });
    const dragRef = useRef({ isDragging: false, startX: 0, startY: 0, startPanX: 0, startPanY: 0 });

    // Calculate distance between two touch points
    const getTouchDistance = (touches) => {
        const dx = touches[0].clientX - touches[1].clientX;
        const dy = touches[0].clientY - touches[1].clientY;
        return Math.sqrt(dx * dx + dy * dy);
    };

    // Mouse down handler for dragging
    const handlePreviewMouseDown = (e) => {
        if (e.button === 0) { // Left mouse button
            dragRef.current = {
                isDragging: true,
                startX: e.clientX,
                startY: e.clientY,
                startPanX: previewState.panX,
                startPanY: previewState.panY
            };
            e.preventDefault();
        }
    };

    // Mouse move handler for dragging
    const handlePreviewMouseMove = (e) => {
        if (dragRef.current.isDragging) {
            const dx = e.clientX - dragRef.current.startX;
            const dy = e.clientY - dragRef.current.startY;
            setPreviewState(prev => ({
                ...prev,
                panX: dragRef.current.startPanX + dx,
                panY: dragRef.current.startPanY + dy
            }));
        }
    };

    // Mouse up handler for dragging
    const handlePreviewMouseUp = () => {
        dragRef.current.isDragging = false;
    };

    // Touch start handler for pinch-to-zoom and pan
    const handlePreviewTouchStart = (e) => {
        if (e.touches.length === 2) {
            // Pinch gesture
            const dist = getTouchDistance(e.touches);
            const rect = e.currentTarget.getBoundingClientRect();
            const cx = rect.left + rect.width / 2;
            const cy = rect.top + rect.height / 2;

            const midX = (e.touches[0].clientX + e.touches[1].clientX) / 2;
            const midY = (e.touches[0].clientY + e.touches[1].clientY) / 2;

            pinchRef.current = {
                initialDistance: dist,
                initialZoom: previewState.zoom,
                initialPanX: previewState.panX,
                initialPanY: previewState.panY,
                startCenterX: midX - cx,
                startCenterY: midY - cy
            };
            dragRef.current.isDragging = false;
        } else if (e.touches.length === 1) {
            // Single finger drag
            dragRef.current = {
                isDragging: true,
                startX: e.touches[0].clientX,
                startY: e.touches[0].clientY,
                startPanX: previewState.panX,
                startPanY: previewState.panY
            };
        }
    };

    // Touch move handler for pinch-to-zoom and pan
    const handlePreviewTouchMove = (e) => {
        if (e.touches.length === 2) {
            // Pinch gesture
            const currentDistance = getTouchDistance(e.touches);
            const rect = e.currentTarget.getBoundingClientRect();
            const cx = rect.left + rect.width / 2;
            const cy = rect.top + rect.height / 2;

            const midX = (e.touches[0].clientX + e.touches[1].clientX) / 2;
            const midY = (e.touches[0].clientY + e.touches[1].clientY) / 2;

            const currentCenterX = midX - cx;
            const currentCenterY = midY - cy;

            const scale = currentDistance / pinchRef.current.initialDistance;
            let newZoom = pinchRef.current.initialZoom * scale;
            newZoom = Math.min(Math.max(newZoom, 0.5), 3.0); // Clamp

            // Calculate new pan to keep the focal point under the center of the pinch
            const effectiveScale = newZoom / pinchRef.current.initialZoom;
            const newPanX = currentCenterX - (pinchRef.current.startCenterX - pinchRef.current.initialPanX) * effectiveScale;
            const newPanY = currentCenterY - (pinchRef.current.startCenterY - pinchRef.current.initialPanY) * effectiveScale;

            setPreviewState(prev => ({
                ...prev,
                zoom: newZoom,
                panX: newPanX,
                panY: newPanY
            }));
        } else if (e.touches.length === 1 && dragRef.current.isDragging) {
            // Single finger pan
            const dx = e.touches[0].clientX - dragRef.current.startX;
            const dy = e.touches[0].clientY - dragRef.current.startY;
            setPreviewState(prev => ({
                ...prev,
                panX: dragRef.current.startPanX + dx,
                panY: dragRef.current.startPanY + dy
            }));
        }
    };

    // Touch end handler
    const handlePreviewTouchEnd = () => {
        dragRef.current.isDragging = false;
    };

    // ============== Drag & Drop Handlers (PC Only) ==============
    const handleDragEnter = (e) => {
        e.preventDefault();
        e.stopPropagation();
        dragCounter.current++;
        if (e.dataTransfer.items && e.dataTransfer.items.length > 0) {
            setIsDragging(true);
        }
    };

    const handleDragLeave = (e) => {
        e.preventDefault();
        e.stopPropagation();
        dragCounter.current--;
        if (dragCounter.current === 0) {
            setIsDragging(false);
        }
    };

    const handleDragOver = (e) => {
        e.preventDefault();
        e.stopPropagation();
    };

    const handleFileDrop = async (e) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragging(false);
        dragCounter.current = 0;

        // Only handle on PC (Electron)
        if (!isElectron()) return;

        const files = e.dataTransfer.files;
        if (!files || files.length === 0) return;

        const finalImages = [];
        let hasConverted = false;

        setIsExporting(true);
        setExportFormat('docx-conversion');

        try {
            for (const file of files) {
                const name = file.name.toLowerCase();

                // Read file as ArrayBuffer
                const arrayBuffer = await file.arrayBuffer();
                const bytes = new Uint8Array(arrayBuffer);

                // Handle Word Documents, PowerPoint, and PDFs
                if (name.endsWith('.docx') || name.endsWith('.doc') || name.endsWith('.pptx') || name.endsWith('.ppt') || name.endsWith('.pdf')) {
                    hasConverted = true;

                    try {
                        // Call Electron API directly with file buffer and extension info
                        const result = await window.electronAPI.convertDocxToImageWord(
                            Array.from(bytes),
                            false, // combinePages = false for separate images
                            file.name // Pass filename so backend can detect PDF vs DOCX
                        );

                        if (result.success && result.images) {
                            for (const imgData of result.images) {
                                // imgData has buffer property (array of bytes)
                                const imgBuffer = imgData.buffer || imgData;
                                let binary = '';
                                const len = imgBuffer.length;
                                for (let i = 0; i < len; i++) {
                                    binary += String.fromCharCode(imgBuffer[i]);
                                }
                                finalImages.push(`data:image/png;base64,${btoa(binary)}`);
                            }
                        } else {
                            console.error('Conversion failed:', result.error || 'Unknown error');
                            showToast(`فشل تحويل: ${file.name}`, 'error');
                        }
                    } catch (convErr) {
                        console.error('Conversion error for', file.name, ':', convErr);
                        showToast(`فشل تحويل: ${file.name}`, 'error');
                    }
                } else if (file.type.startsWith('image/')) {
                    // Handle images directly
                    const base64 = await new Promise((resolve) => {
                        const reader = new FileReader();
                        reader.onload = () => resolve(reader.result);
                        reader.readAsDataURL(file);
                    });
                    finalImages.push(base64);
                }
            }

            if (finalImages.length > 0) {
                // If preview is already open, append to existing images (like "اضافة مستندات" button)
                if (previewState.show) {
                    setPreviewState(prev => ({
                        ...prev,
                        images: [...prev.images, ...finalImages]
                    }));
                    setAddedImages(prev => [...prev, ...finalImages]); // Track for export
                    showToast(`تم إضافة ${finalImages.length} ${finalImages.length > 1 ? 'صفحات' : 'صفحة'} بنجاح`);
                } else {
                    // Open new preview with dropped files
                    setPreviewState({
                        show: true,
                        images: finalImages,
                        currentPage: 0,
                        zoom: 1.0,
                        panX: 0,
                        panY: 0,
                        isLoading: false,
                        isExternal: true
                    });
                    if (hasConverted) {
                        showToast('تم تحويل الملفات بنجاح', 'success');
                    } else {
                        showToast(`تم استلام ${finalImages.length > 1 ? finalImages.length + ' صور' : 'صورة'}`, 'success');
                    }
                }
            }
        } catch (error) {
            console.error('Error processing dropped files:', error);
            showToast('حدث خطأ أثناء معالجة الملفات', 'error');
        } finally {
            setIsExporting(false);
        }
    };

    return (
        <div
            className={`app ${isDragging ? 'dragging' : ''}`}
            onDragEnter={handleDragEnter}
            onDragLeave={handleDragLeave}
            onDragOver={handleDragOver}
            onDrop={handleFileDrop}
        >


            {/* Custom Confirmation Modal */}
            {confirmModal.show && (
                <div className="loading-overlay">
                    <div className="confirm-content">
                        <span className="confirm-message">{confirmModal.message}</span>
                        <div className="confirm-buttons">
                            <button
                                className="confirm-btn confirm"
                                onClick={confirmModal.onConfirm}
                            >
                                نعم
                            </button>
                            <button
                                className="confirm-btn cancel"
                                onClick={() => setConfirmModal({ show: false, message: '', onConfirm: null })}
                            >
                                إلغاء
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {/* Preview Overlay */}
            {previewState.show && (
                <div className="preview-overlay">
                    {previewState.isLoading ? (
                        <div className="loading-content">
                            <div className="loading-spinner"></div>
                            <span className="loading-text">جاري إنشاء المعاينة...</span>
                            <button
                                className="cancel-export-btn"
                                onClick={closePreview}
                            >
                                إلغاء
                            </button>
                        </div>
                    ) : (
                        <div className="preview-modal-container">
                            {/* Top Controls Bar */}
                            <div className="preview-top-bar">
                                <button className="preview-control-btn close-btn" onClick={closePreview} title="إغلاق">
                                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                        <line x1="18" y1="6" x2="6" y2="18" />
                                        <line x1="6" y1="6" x2="18" y2="18" />
                                    </svg>
                                </button>
                                <div className="preview-zoom-controls">
                                    <button className="preview-control-btn" onClick={previewZoomOut} title="تصغير">
                                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                            <circle cx="11" cy="11" r="8" />
                                            <line x1="7" y1="11" x2="15" y2="11" />
                                        </svg>
                                    </button>
                                    <span className="preview-zoom-level">{Math.round(previewState.zoom * 100)}%</span>
                                    <button className="preview-control-btn" onClick={previewZoomIn} title="تكبير">
                                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                            <circle cx="11" cy="11" r="8" />
                                            <line x1="7" y1="11" x2="15" y2="11" />
                                            <line x1="11" y1="7" x2="11" y2="15" />
                                        </svg>
                                    </button>
                                </div>
                                <div className="preview-page-indicator">
                                    {previewState.images.length > 0 && (
                                        <span>{previewState.currentPage + 1} / {previewState.images.length}</span>
                                    )}
                                </div>
                            </div>

                            {/* Image Container */}
                            <div
                                className="preview-image-container"
                                onWheel={handlePreviewWheel}
                                onMouseDown={handlePreviewMouseDown}
                                onMouseMove={handlePreviewMouseMove}
                                onMouseUp={handlePreviewMouseUp}
                                onMouseLeave={handlePreviewMouseUp}
                                onTouchStart={handlePreviewTouchStart}
                                onTouchMove={handlePreviewTouchMove}
                                onTouchEnd={handlePreviewTouchEnd}
                            >
                                {previewState.images.length > 0 && (
                                    <>
                                        <img
                                            src={previewState.images[previewState.currentPage]}
                                            alt={`صفحة ${previewState.currentPage + 1}`}
                                            className="preview-image"
                                            style={{
                                                transform: `translate(${previewState.panX}px, ${previewState.panY}px) scale(${previewState.zoom})`
                                            }}
                                            draggable={false}
                                        />
                                        {/* Delete button - for added images and external/shared images */}
                                        {(previewState.isExternal || addedImages.includes(previewState.images[previewState.currentPage])) && (
                                            <button
                                                className="preview-delete-image-btn"
                                                onClick={handleDeleteCurrentImage}
                                                title="حذف الصورة"
                                            >
                                                🗑️
                                            </button>
                                        )}
                                    </>
                                )}
                            </div>

                            {/* Navigation Arrows */}
                            {previewState.images.length > 1 && (
                                <>
                                    <button
                                        className="preview-nav-btn prev-btn"
                                        onClick={previewPrevPage}
                                        disabled={previewState.currentPage === 0}
                                        title="الصفحة السابقة"
                                    >
                                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                            <polyline points="9 18 15 12 9 6" />
                                        </svg>
                                    </button>
                                    <button
                                        className="preview-nav-btn next-btn"
                                        onClick={previewNextPage}
                                        disabled={previewState.currentPage === previewState.images.length - 1}
                                        title="الصفحة التالية"
                                    >
                                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                                            <polyline points="15 18 9 12 15 6" />
                                        </svg>
                                    </button>
                                </>
                            )}

                            {/* Bottom Export Bar */}
                            {!previewState.isExternal ? (
                                <div className="preview-export-bar">
                                    <button
                                        className="preview-export-btn docx"
                                        onClick={() => { setExportFormat('docx'); exportDocument('docx'); }}
                                        disabled={isExporting}
                                    >
                                        <span className="icon">📄</span>
                                        <span>Word</span>
                                    </button>
                                    <button
                                        className="preview-export-btn pdf"
                                        onClick={() => { setExportFormat('pdf'); exportDocument('pdf'); }}
                                        disabled={isExporting}
                                    >
                                        <span className="icon">📕</span>
                                        <span>PDF</span>
                                    </button>
                                    <button
                                        className="preview-export-btn image"
                                        onClick={() => { setExportFormat('png'); exportDocument('png'); }}
                                        disabled={isExporting}
                                    >
                                        <span className="icon">🖼️</span>
                                        <span>صورة</span>
                                    </button>
                                    <label className="preview-combine-option">
                                        <input
                                            type="checkbox"
                                            checked={combineImagePages}
                                            onChange={(e) => setCombineImagePages(e.target.checked)}
                                        />
                                        <span>دمج الصفحات</span>
                                    </label>
                                    <button
                                        className="preview-export-btn add-images"
                                        onClick={handleExternalAddImages}
                                        disabled={isExporting}
                                    >
                                        <span className="icon">➕</span>
                                        <span>إضافة مستندات</span>
                                    </button>
                                </div>
                            ) : (
                                <div className="preview-export-bar external-mode">
                                    <button
                                        className="preview-export-btn add-images"
                                        onClick={handleExternalAddImages}
                                        disabled={isExporting}
                                    >
                                        <span className="icon">➕</span>
                                        <span>إضافة مستندات</span>
                                    </button>
                                    <button
                                        className="preview-export-btn export-pdf"
                                        onClick={handleExternalExportPdf}
                                        disabled={isExporting}
                                    >
                                        <span className="icon">📕</span>
                                        <span>تصدير PDF</span>
                                    </button>
                                </div>
                            )}
                        </div>
                    )}
                </div>
            )}

            {/* iOS-style Toast Notification */}
            {toast.show && (
                <div className={`toast-notification ${toast.type}`}>
                    <span className="toast-icon">✓</span>
                    <span className="toast-message">{toast.message}</span>
                </div>
            )}

            <main className="main">
                <div className="form-container">
                    {/* Logo Section */}
                    <div className="form-section logo-section">
                        <h2>الشعار</h2>
                        <div className="logo-upload" onClick={handleLogoSelect}>
                            {formData.logoBase64 && (
                                <button
                                    className="logo-clear-btn"
                                    onClick={handleLogoClear}
                                    title="إزالة الشعار"
                                >
                                    ✕
                                </button>
                            )}
                            {formData.logoBase64 ? (
                                <img
                                    src={`data:image/png;base64,${formData.logoBase64}`}
                                    alt="Logo"
                                    className="logo-preview"
                                />
                            ) : (
                                <div className="logo-placeholder">
                                    <span className="icon">🖼️</span>
                                    <span>اضغط لاختيار الشعار</span>
                                </div>
                            )}
                        </div>
                        {formData.logoName && <p className="logo-name">{formData.logoName}</p>}
                    </div>

                    {/* Header Fields */}
                    <div className="form-section">
                        <h2>معلومات المرسل</h2>
                        <div className="form-row">
                            <div className="form-group">
                                <label htmlFor="parent_company">الشركة الأم:</label>
                                <input
                                    type="text"
                                    id="parent_company"
                                    name="parent_company"
                                    value={formData.parent_company}
                                    onChange={handleChange}
                                    placeholder="الشركة الأم"
                                    className="persistent"
                                />
                            </div>
                            <div className="form-group">
                                <label htmlFor="subsidiary_company">الشركة الفرعية:</label>
                                <input
                                    type="text"
                                    id="subsidiary_company"
                                    name="subsidiary_company"
                                    value={formData.subsidiary_company}
                                    onChange={handleChange}
                                    placeholder="الشركة الفرعية"
                                    className="persistent"
                                />
                            </div>
                            <div className="form-group">
                                <label htmlFor="from">الوحدة المحددة:</label>
                                <input
                                    type="text"
                                    id="from"
                                    name="from"
                                    value={formData.from}
                                    onChange={handleChange}
                                    placeholder="اسم الجهة المرسلة"
                                    className="persistent"
                                />
                            </div>
                        </div>
                    </div>

                    {/* Recipient Fields */}
                    <div className="form-section">
                        <h2>معلومات المستلم</h2>
                        <div className="form-row">
                            <div className="form-group">
                                <label htmlFor="to">إلى (المستلم):</label>
                                <input
                                    type="text"
                                    id="to"
                                    name="to"
                                    value={formData.to}
                                    onChange={handleChange}
                                    placeholder="اسم المستلم"
                                    className="prefilled"
                                />
                            </div>
                            <div className="form-group">
                                <label htmlFor="to_the">&nbsp;</label>
                                <input
                                    type="text"
                                    id="to_the"
                                    name="to_the"
                                    value={formData.to_the}
                                    onChange={handleChange}
                                    placeholder="منصب المستلم"
                                    className="prefilled"
                                />
                            </div>
                        </div>
                    </div>

                    {/* Content Fields */}
                    <div className="form-section">
                        <h2>محتوى الرسالة</h2>
                        <div className="form-row">
                            <div className="form-group">
                                <label htmlFor="greetings">التحية:</label>
                                <input
                                    type="text"
                                    id="greetings"
                                    name="greetings"
                                    value={formData.greetings}
                                    onChange={handleChange}
                                    placeholder="مثال: السلام عليكم ورحمة الله وبركاته"
                                    className="persistent"
                                />
                            </div>
                        </div>
                        <div className="form-row">
                            <div className="form-group">
                                <label htmlFor="subject_name">عنوان الموضوع:</label>
                                <input
                                    type="text"
                                    id="subject_name"
                                    name="subject_name"
                                    value={formData.subject_name}
                                    onChange={handleChange}
                                    placeholder="عنوان الموضوع"
                                    className="prefilled"
                                />
                            </div>
                        </div>
                        <div className="form-row">
                            <div className="form-group full-width">
                                <label htmlFor="subject">الموضوع:</label>
                                <textarea
                                    id="subject"
                                    name="subject"
                                    value={formData.subject}
                                    onChange={handleChange}
                                    placeholder="نص الرسالة..."
                                    rows={6}
                                    className="prefilled"
                                />
                                <div style={{ marginTop: '10px' }}>
                                    <label className="checkbox-label">
                                        <input
                                            type="checkbox"
                                            name="useTable"
                                            checked={formData.useTable || false}
                                            onChange={handleChange}
                                        />
                                        <span className="checkmark"></span>
                                        <span>استخدام جدول</span>
                                    </label>
                                </div>

                                {formData.useTable && (
                                    <div className="table-editor-container">
                                        <div className="table-wrapper">
                                            <table className="editor-table">
                                                <tbody>
                                                    {formData.tableData && formData.tableData.map((row, rowIndex) => (
                                                        <tr key={rowIndex}>
                                                            {row.map((cell, colIndex) => (
                                                                <td key={colIndex}>
                                                                    <div className="cell-wrapper">
                                                                        <input
                                                                            type="text"
                                                                            value={cell}
                                                                            onChange={(e) => handleTableChange(rowIndex, colIndex, e.target.value)}
                                                                            placeholder={`Cell ${rowIndex + 1}-${colIndex + 1}`}
                                                                        />
                                                                        {/* Only show delete controls for actual content rows/cols to avoid emptying table completely easily */}
                                                                        {rowIndex > 0 && colIndex === row.length - 1 && (
                                                                            <button type="button" className="delete-row-btn" onClick={() => removeTableRow(rowIndex)} title="حذف الصف">×</button>
                                                                        )}
                                                                        {rowIndex === 0 && colIndex > 0 && (
                                                                            <button type="button" className="delete-col-btn" onClick={() => removeTableColumn(colIndex)} title="حذف العمود">×</button>
                                                                        )}
                                                                        {rowIndex === 0 && (
                                                                            <div className="col-adder" onClick={() => insertTableColumn(colIndex)} title="إضافة عمود جديد" />
                                                                        )}
                                                                        {colIndex === 0 && (
                                                                            <div className="row-adder" onClick={() => insertTableRow(rowIndex)} title="إضافة صف جديد" />
                                                                        )}
                                                                    </div>
                                                                </td>
                                                            ))}
                                                        </tr>
                                                    ))}
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                )}
                            </div>
                        </div>
                    </div>

                    {/* Closing Fields */}
                    <div className="form-section">
                        <h2>الختام</h2>
                        <div className="form-row">
                            <div className="form-group">
                                <label htmlFor="ending">الخاتمة:</label>
                                <input
                                    type="text"
                                    id="ending"
                                    name="ending"
                                    value={formData.ending}
                                    onChange={handleChange}
                                    placeholder="مثال: وتقبلوا فائق الاحترام والتقدير"
                                    className="persistent"
                                />
                            </div>
                            <div className="form-group">
                                <label htmlFor="sign">التوقيع:</label>
                                <textarea
                                    id="sign"
                                    name="sign"
                                    value={formData.sign}
                                    onChange={handleChange}
                                    placeholder="اسم الموقع"
                                    className="persistent"
                                    rows={2}
                                />
                            </div>
                        </div>
                        <div className="form-row">
                            <div className="form-group">
                                <label htmlFor="copy_to">نسخة إلى:</label>
                                <textarea
                                    id="copy_to"
                                    name="copy_to"
                                    value={formData.copy_to}
                                    onChange={handleChange}
                                    placeholder="الجهات المنسوخ إليها (اضغط Enter لسطر جديد)"
                                    className="persistent"
                                    rows={3}
                                />
                            </div>
                        </div>
                    </div>

                    {/* Date Toggle */}
                    <div className="form-section">
                        <h2>إعدادات التاريخ</h2>
                        <div className="form-row">
                            <label className="checkbox-label">
                                <input
                                    type="checkbox"
                                    name="showDate"
                                    checked={formData.showDate}
                                    onChange={handleChange}
                                />
                                <span className="checkmark"></span>
                                <span>إظهار مربع التاريخ (التاريخ التلقائي)</span>
                            </label>
                        </div>
                    </div>

                    {/* Export Buttons */}
                    <div className="form-section export-section">
                        <h2>تصدير المستند</h2>
                        <div className="export-buttons">
                            <button className="export-btn docx" onClick={() => { setExportFormat('docx'); exportDocument('docx'); }} disabled={isExporting}>
                                <span className="icon">📄</span>
                                <span>Word</span>
                            </button>
                            <button className="export-btn pdf" onClick={() => { setExportFormat('pdf'); exportDocument('pdf'); }} disabled={isExporting}>
                                <span className="icon">📕</span>
                                <span>PDF</span>
                            </button>
                            <button className="export-btn image" onClick={() => { setExportFormat('png'); exportDocument('png'); }} disabled={isExporting}>
                                <span className="icon">🖼️</span>
                                <span>صورة</span>
                            </button>
                            <label className="combine-checkbox">
                                <input
                                    type="checkbox"
                                    checked={combineImagePages}
                                    onChange={(e) => setCombineImagePages(e.target.checked)}
                                    disabled={isExporting}
                                />
                                <span className="checkmark-small"></span>
                                <span className="combine-label">دمج كل الصفحات في صورة واحدة</span>
                            </label>
                        </div>
                    </div>
                </div>
            </main>

            {/* Footer */}
            <footer className="footer">
                <p>
                    تم التطوير بواسطة{' '}
                    <a
                        href="#"
                        onClick={(e) => { e.preventDefault(); openExternalUrl('https://www.linkedin.com/in/moh-d-m4x/'); }}
                    >
                        محمد عبدالله
                    </a>
                    {' '}|{' '}
                    الكود المصدري على{' '}
                    <a
                        href="#"
                        onClick={(e) => { e.preventDefault(); openExternalUrl('https://github.com/moh-d-m4x/AutoWriter'); }}
                    >
                        GitHub
                    </a>
                </p>
            </footer>

            {/* Floating Preview Button */}
            <button className="floating-preview-btn" title="معاينة المستند" onClick={handlePreview}>
                <svg className="preview-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
                    <circle cx="12" cy="12" r="3" />
                </svg>
            </button>

            {/* Loading Overlay with Growing Progress Arc */}
            {isExporting && (
                <div className="loading-overlay">
                    <div className="loading-content">
                        <div className="loading-progress-container">
                            {/* SVG Spinner with growing arc */}
                            <svg className="loading-progress-ring" width="50" height="50" viewBox="0 0 50 50">
                                {/* Background circle (faint) */}
                                <circle
                                    cx="25"
                                    cy="25"
                                    r="20"
                                    fill="none"
                                    stroke="rgba(255,255,255,0.15)"
                                    strokeWidth="3"
                                />
                                {/* Progress arc - starts small, grows to full circle */}
                                <circle
                                    className="loading-progress-arc"
                                    cx="25"
                                    cy="25"
                                    r="20"
                                    fill="none"
                                    stroke="white"
                                    strokeWidth="3"
                                    strokeLinecap="round"
                                    strokeDasharray={`${Math.max(10, (exportProgress / 100) * 125.66)} 125.66`}
                                    transform="rotate(-90 25 25)"
                                />
                            </svg>
                            {/* Percentage text in center */}
                            {exportProgress > 0 && (
                                <span className="loading-progress-text">{exportProgress}%</span>
                            )}
                        </div>
                        <p className="loading-text">
                            {exportFormat === 'docx-conversion' ? 'جاري فتح الملف...' : 'جاري التصدير...'}
                        </p>
                        {/* File counter for multiple files */}
                        {fileProgress.total > 1 && (
                            <p className="loading-file-counter">
                                الملف {fileProgress.current} من {fileProgress.total}
                            </p>
                        )}
                        {/* Cancel Button - Hidden for DOCX and Conversion */}
                        {exportFormat !== 'docx' && exportFormat !== 'docx-conversion' && (
                            <button
                                className="cancel-export-btn"
                                onClick={async () => {
                                    exportCancelledRef.current = true;
                                    if (isElectron() && window.electronAPI) {
                                        await window.electronAPI.cancelExport();
                                    }
                                    setIsExporting(false);
                                    setExportProgress(0);
                                    showToast('تم إلغاء التصدير', 'info');
                                }}
                            >
                                إلغاء
                            </button>
                        )}
                    </div>
                </div>
            )}

            {/* Source Selection Dialog (Android only) */}
            {showSourceDialog && (
                <div className="source-dialog-overlay" onClick={() => setShowSourceDialog(false)}>
                    <div className="source-dialog" onClick={(e) => e.stopPropagation()}>
                        <h3>إضافة مستندات</h3>
                        <div className="source-dialog-options">
                            <button className="source-option" onClick={handleAndroidAddFromCamera}>
                                <span className="source-icon">📷</span>
                                <span>الكاميرا</span>
                            </button>
                            <button className="source-option" onClick={handleAndroidAddFromFilePicker}>
                                <span className="source-icon">📁</span>
                                <span>الملفات</span>
                            </button>
                        </div>
                        <button className="source-cancel-btn" onClick={() => setShowSourceDialog(false)}>
                            إلغاء
                        </button>
                    </div>
                </div>
            )}
        </div>
    )
}

export default App
