/**
 * ===========================================
 * LibreOfficeKit Capacitor Plugin
 * ===========================================
 * 
 * LibreOfficeKit code implementation
 * 
 * This plugin provides document conversion functionality on Android
 * using LibreOfficeKit (LOK) for DOCX to PDF conversion.
 * 
 * NOTE: LibreOffice initialization is LAZY - it only happens when
 * conversion is first requested to avoid crash on app startup.
 * 
 * @author AutoWriter
 * @version 2.0.0
 */
package com.autowriter.app;

import android.util.Base64;
import android.util.Log;
import android.os.Handler;
import android.os.Looper;

import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * LibreOfficeKit code implementation
 * 
 * LOKPlugin - Capacitor bridge for LibreOfficeKit document conversion
 * NOTE: LOK is NOT initialized at startup to prevent crashes
 */
@CapacitorPlugin(name = "LOKPlugin")
public class LOKPlugin extends Plugin {

    private static final String TAG = "LOKPlugin";

    // LibreOfficeKit code implementation - Executor for background tasks
    private final ExecutorService executor = Executors.newSingleThreadExecutor();

    // LibreOfficeKit code implementation - Converter instance (lazy init)
    private LOKConverter converter = null;

    // Track if we've tried to initialize
    private boolean initAttempted = false;
    private String initError = null;

    // Shared Image Logic
    // Shared Image Logic
    public static java.util.List<String> pendingSharedImagePaths = java.util.Collections
            .synchronizedList(new java.util.ArrayList<>());

    /**
     * LibreOfficeKit code implementation
     * Initialize the plugin - but NOT LibreOffice (lazy init)
     */
    @Override
    public void load() {
        super.load();
        Log.d(TAG, "LOKPlugin: Loaded (LOK init will be lazy)");
        // DO NOT initialize LibreOffice here - it causes native crashes
        // Initialization happens on first conversion request
    }

    /**
     * Lazy initialize LibreOffice when first needed
     */
    private synchronized void ensureInitialized() {
        if (initAttempted) {
            return; // Already tried (success or failure)
        }
        initAttempted = true;

        Log.d(TAG, "LOKPlugin: Lazy-initializing LibreOfficeKit...");
        try {
            converter = new LOKConverter(getContext());
            if (converter.isReady()) {
                Log.d(TAG, "LOKPlugin: LibreOfficeKit initialized successfully");
            } else {
                initError = "LOK initialization returned not ready";
                Log.e(TAG, "LOKPlugin: " + initError);
            }
        } catch (UnsatisfiedLinkError e) {
            initError = "Native library error: " + e.getMessage();
            Log.e(TAG, "LOKPlugin: " + initError, e);
            converter = null;
        } catch (Exception e) {
            initError = "Init error: " + e.getMessage();
            Log.e(TAG, "LOKPlugin: " + initError, e);
            converter = null;
        }
    }

    /**
     * LibreOfficeKit code implementation
     * Convert DOCX document to PDF
     */
    @PluginMethod()
    public void convertToPdf(PluginCall call) {
        Log.d(TAG, "LOKPlugin: convertToPdf called");

        String docxBase64 = call.getString("docxBuffer");

        if (docxBase64 == null || docxBase64.isEmpty()) {
            JSObject result = new JSObject();
            result.put("success", false);
            result.put("error", "No document data provided");
            call.resolve(result);
            return;
        }

        // LibreOfficeKit code implementation - Run conversion in background
        executor.execute(() -> {
            try {
                // Lazy init on first use
                ensureInitialized();

                if (converter == null || !converter.isReady()) {
                    throw new Exception("LibreOffice not available: " +
                            (initError != null ? initError : "initialization failed"));
                }

                byte[] docxBytes = Base64.decode(docxBase64, Base64.DEFAULT);

                File tempDir = getContext().getCacheDir();
                File inputFile = new File(tempDir, "input_" + System.currentTimeMillis() + ".docx");
                File outputFile = new File(tempDir, "output_" + System.currentTimeMillis() + ".pdf");

                try (FileOutputStream fos = new FileOutputStream(inputFile)) {
                    fos.write(docxBytes);
                }

                Log.d(TAG, "LOKPlugin: Converting " + inputFile.getAbsolutePath() + " to PDF");

                // LibreOfficeKit code implementation - Perform conversion
                boolean conversionSuccess = converter.convertToPdf(
                        inputFile.getAbsolutePath(),
                        outputFile.getAbsolutePath());

                if (conversionSuccess && outputFile.exists()) {
                    byte[] pdfBytes = readFileToBytes(outputFile);
                    String pdfBase64 = Base64.encodeToString(pdfBytes, Base64.NO_WRAP);

                    inputFile.delete();
                    outputFile.delete();

                    JSObject result = new JSObject();
                    result.put("success", true);
                    result.put("buffer", pdfBase64);

                    new Handler(Looper.getMainLooper()).post(() -> call.resolve(result));
                    Log.d(TAG, "LOKPlugin: PDF conversion successful");
                } else {
                    throw new Exception("Conversion failed - output file not created");
                }

            } catch (Exception e) {
                Log.e(TAG, "LOKPlugin: Conversion error", e);

                JSObject result = new JSObject();
                result.put("success", false);
                result.put("error", e.getMessage());

                new Handler(Looper.getMainLooper()).post(() -> call.resolve(result));
            }
        });
    }

    /**
     * LibreOfficeKit code implementation
     * Convert DOCX document to PNG Image
     */
    @PluginMethod()
    public void convertToImage(PluginCall call) {
        Log.d(TAG, "LOKPlugin: convertToImage called");

        String docxBase64 = call.getString("docxBuffer");

        if (docxBase64 == null || docxBase64.isEmpty()) {
            JSObject result = new JSObject();
            result.put("success", false);
            result.put("error", "No document data provided");
            call.resolve(result);
            return;
        }

        // Run conversion in background
        executor.execute(() -> {
            try {
                // Lazy init on first use
                ensureInitialized();

                if (converter == null || !converter.isReady()) {
                    throw new Exception(
                            "LibreOfficeKit not initialized: " + (initError != null ? initError : "Unknown error"));
                }

                // Decode base64 to bytes
                byte[] docxBytes = Base64.decode(docxBase64, Base64.DEFAULT);
                String extension = call.getString("extension", ".docx");
                if (!extension.startsWith("."))
                    extension = "." + extension;

                // Create temp files
                File cacheDir = getContext().getCacheDir();
                File inputFile = new File(cacheDir, "input_" + System.currentTimeMillis() + extension);
                File outputFile = new File(cacheDir, "output_" + System.currentTimeMillis() + ".png");

                // Write Buffer to temp file
                try (FileOutputStream fos = new FileOutputStream(inputFile)) {
                    fos.write(docxBytes);
                }

                // Perform conversion via PDF (DOCX → PDF → PNG)
                boolean conversionSuccess = converter.convertToImageViaPdf(
                        inputFile.getAbsolutePath(),
                        outputFile.getAbsolutePath(),
                        true); // Combine all pages

                if (conversionSuccess && outputFile.exists()) {
                    byte[] pngBytes = readFileToBytes(outputFile);
                    String pngBase64 = Base64.encodeToString(pngBytes, Base64.NO_WRAP);

                    inputFile.delete();
                    outputFile.delete();

                    JSObject result = new JSObject();
                    result.put("success", true);
                    result.put("buffer", pngBase64);

                    new Handler(Looper.getMainLooper()).post(() -> call.resolve(result));
                    Log.d(TAG, "LOKPlugin: Image conversion successful");
                } else {
                    throw new Exception("Image conversion failed - output file not created");
                }

            } catch (Exception e) {
                Log.e(TAG, "LOKPlugin: Image conversion error", e);

                JSObject result = new JSObject();
                result.put("success", false);
                result.put("error", e.getMessage());

                new Handler(Looper.getMainLooper()).post(() -> call.resolve(result));
            }
        });
    }

    /**
     * LibreOfficeKit code implementation
     * Convert DOCX document to multiple PNG Images (separate pages)
     */
    @PluginMethod()
    public void convertToImages(PluginCall call) {
        Log.d(TAG, "LOKPlugin: convertToImages called");

        String docxBase64 = call.getString("docxBuffer");
        String baseName = call.getString("baseName", "page");

        if (docxBase64 == null || docxBase64.isEmpty()) {
            JSObject result = new JSObject();
            result.put("success", false);
            result.put("error", "No document data provided");
            call.resolve(result);
            return;
        }

        // Run conversion in background
        executor.execute(() -> {
            try {
                // Lazy init on first use
                ensureInitialized();

                if (converter == null || !converter.isReady()) {
                    throw new Exception(
                            "LibreOfficeKit not initialized: " + (initError != null ? initError : "Unknown error"));
                }

                // Decode base64 to bytes
                byte[] docxBytes = Base64.decode(docxBase64, Base64.DEFAULT);
                String extension = call.getString("extension", ".docx");
                if (!extension.startsWith("."))
                    extension = "." + extension;

                // Create temp files
                File cacheDir = getContext().getCacheDir();
                File inputFile = new File(cacheDir, "input_" + System.currentTimeMillis() + extension);

                // Write Buffer to temp file
                try (FileOutputStream fos = new FileOutputStream(inputFile)) {
                    fos.write(docxBytes);
                }

                // Perform conversion via PDF (DOCX → PDF → PNG)
                String[] outputPaths = converter.convertToImagesViaPdf(
                        inputFile.getAbsolutePath(),
                        cacheDir.getAbsolutePath(),
                        baseName);

                inputFile.delete();

                if (outputPaths != null && outputPaths.length > 0) {
                    // Read all images and convert to base64
                    org.json.JSONArray imagesArray = new org.json.JSONArray();

                    for (String path : outputPaths) {
                        File imageFile = new File(path);
                        if (imageFile.exists()) {
                            byte[] imageBytes = readFileToBytes(imageFile);
                            String imageBase64 = Base64.encodeToString(imageBytes, Base64.NO_WRAP);
                            imagesArray.put(imageBase64);
                            imageFile.delete();
                        }
                    }

                    JSObject result = new JSObject();
                    result.put("success", true);
                    result.put("count", imagesArray.length());
                    result.put("images", imagesArray);

                    new Handler(Looper.getMainLooper()).post(() -> call.resolve(result));
                    Log.d(TAG,
                            "LOKPlugin: Multi-page image conversion successful - " + imagesArray.length() + " pages");
                } else {
                    throw new Exception("Image conversion failed - no output files created");
                }

            } catch (Exception e) {
                Log.e(TAG, "LOKPlugin: Multi-page image conversion error", e);

                JSObject result = new JSObject();
                result.put("success", false);
                result.put("error", e.getMessage());

                new Handler(Looper.getMainLooper()).post(() -> call.resolve(result));
            }
        });
    }

    /**
     * LibreOfficeKit code implementation
     * Check for any pending shared image (from Intent)
     */
    @PluginMethod()
    public void getSharedImage(PluginCall call) {
        synchronized (pendingSharedImagePaths) {
            if (pendingSharedImagePaths != null && !pendingSharedImagePaths.isEmpty()) {
                try {
                    org.json.JSONArray imagesArray = new org.json.JSONArray();
                    org.json.JSONArray filesArray = new org.json.JSONArray();

                    // Supported file extensions - must match JavaScript filtering
                    java.util.Set<String> supportedExtensions = new java.util.HashSet<>(java.util.Arrays.asList(
                            ".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp",
                            ".pdf", ".docx", ".doc", ".pptx", ".ppt", ".xlsx", ".xls"));

                    for (String path : pendingSharedImagePaths) {
                        File imageFile = new File(path);
                        if (imageFile.exists()) {
                            // Check file extension before loading
                            String fileName = imageFile.getName().toLowerCase();
                            String extension = "";
                            int dotIndex = fileName.lastIndexOf('.');
                            if (dotIndex > 0) {
                                extension = fileName.substring(dotIndex);
                            }

                            // Skip unsupported files to prevent OOM
                            if (!supportedExtensions.contains(extension)) {
                                Log.d(TAG, "Skipping unsupported file: " + fileName);
                                // Clean up the unsupported file
                                imageFile.delete();
                                continue;
                            }

                            byte[] imageBytes = readFileToBytes(imageFile);
                            String imageBase64 = Base64.encodeToString(imageBytes, Base64.NO_WRAP);

                            // Legacy support
                            imagesArray.put(imageBase64);

                            // New generic file support
                            JSObject fileObj = new JSObject();
                            fileObj.put("data", imageBase64);
                            fileObj.put("name", imageFile.getName());
                            filesArray.put(fileObj);

                            if (imageFile.exists()) {
                                imageFile.delete();
                            }
                        }
                    }

                    // Clear pending immediately
                    pendingSharedImagePaths.clear();

                    if (filesArray.length() > 0) {
                        JSObject result = new JSObject();
                        result.put("hasImage", true);
                        result.put("images", imagesArray); // Keep legacy
                        result.put("files", filesArray); // Add new

                        // For backward compatibility
                        try {
                            result.put("image", imagesArray.getString(0));
                        } catch (Exception e) {
                        }

                        call.resolve(result);
                        return;
                    }
                } catch (Exception e) {
                    Log.e(TAG, "Error reading shared image", e);
                }
            }
        }

        JSObject result = new JSObject();
        result.put("hasImage", false);
        call.resolve(result);
    }

    /**
     * LibreOfficeKit code implementation
     * Check if LibreOfficeKit is available and ready
     */
    @PluginMethod()
    public void isReady(PluginCall call) {
        JSObject result = new JSObject();
        // Don't trigger init just for ready check
        result.put("ready", converter != null && converter.isReady());
        result.put("initialized", initAttempted);
        if (initError != null) {
            result.put("error", initError);
        }
        call.resolve(result);
    }

    /**
     * LibreOfficeKit code implementation
     * Helper to read file into byte array
     */
    private byte[] readFileToBytes(File file) throws IOException {
        byte[] bytes = new byte[(int) file.length()];
        try (FileInputStream fis = new FileInputStream(file)) {
            fis.read(bytes);
        }
        return bytes;
    }

    /**
     * LibreOfficeKit code implementation
     * Cleanup when plugin is destroyed
     */
    @Override
    protected void handleOnDestroy() {
        super.handleOnDestroy();
        executor.shutdown();
        if (converter != null) {
            converter.cleanup();
        }
        Log.d(TAG, "LOKPlugin: Cleaned up");
    }

    /**
     * Convert document to image FILES (not Base64)
     * Returns file paths and emits progress events
     */
    @PluginMethod()
    public void convertToImageFiles(PluginCall call) {
        Log.d(TAG, "LOKPlugin: convertToImageFiles called");

        String docxBase64 = call.getString("docxBuffer");
        String baseName = call.getString("baseName", "page");

        if (docxBase64 == null || docxBase64.isEmpty()) {
            JSObject result = new JSObject();
            result.put("success", false);
            result.put("error", "No document data provided");
            call.resolve(result);
            return;
        }

        executor.execute(() -> {
            try {
                ensureInitialized();

                if (converter == null || !converter.isReady()) {
                    throw new Exception("LibreOfficeKit not initialized");
                }

                byte[] docxBytes = Base64.decode(docxBase64, Base64.DEFAULT);
                String extension = call.getString("extension", ".docx");
                if (!extension.startsWith("."))
                    extension = "." + extension;

                File cacheDir = getContext().getCacheDir();
                File inputFile = new File(cacheDir, "input_" + System.currentTimeMillis() + extension);

                try (FileOutputStream fos = new FileOutputStream(inputFile)) {
                    fos.write(docxBytes);
                }

                // Use converter with progress callback
                String[] outputPaths = converter.convertToImagesViaPdfWithProgress(
                        inputFile.getAbsolutePath(),
                        cacheDir.getAbsolutePath(),
                        baseName,
                        (current, total) -> {
                            // Emit progress event to JavaScript
                            JSObject progress = new JSObject();
                            progress.put("current", current);
                            progress.put("total", total);
                            progress.put("percent", (int) ((current * 100.0) / total));
                            new Handler(Looper.getMainLooper())
                                    .post(() -> notifyListeners("conversionProgress", progress));
                        });

                inputFile.delete();

                if (outputPaths != null && outputPaths.length > 0) {
                    org.json.JSONArray pathsArray = new org.json.JSONArray();
                    for (String path : outputPaths) {
                        pathsArray.put(path);
                    }

                    JSObject result = new JSObject();
                    result.put("success", true);
                    result.put("count", outputPaths.length);
                    result.put("paths", pathsArray);

                    new Handler(Looper.getMainLooper()).post(() -> call.resolve(result));
                    Log.d(TAG, "LOKPlugin: convertToImageFiles successful - " + outputPaths.length + " files");
                } else {
                    throw new Exception("Conversion failed - no output files");
                }

            } catch (Exception e) {
                Log.e(TAG, "LOKPlugin: convertToImageFiles error", e);
                JSObject result = new JSObject();
                result.put("success", false);
                result.put("error", e.getMessage());
                new Handler(Looper.getMainLooper()).post(() -> call.resolve(result));
            }
        });
    }

    /**
     * Get single image as Base64 from file path
     */
    @PluginMethod()
    public void getImageAsBase64(PluginCall call) {
        String path = call.getString("path");

        if (path == null || path.isEmpty()) {
            JSObject result = new JSObject();
            result.put("success", false);
            result.put("error", "No path provided");
            call.resolve(result);
            return;
        }

        executor.execute(() -> {
            try {
                File imageFile = new File(path);
                if (!imageFile.exists()) {
                    throw new Exception("File not found: " + path);
                }

                byte[] imageBytes = readFileToBytes(imageFile);
                String base64 = Base64.encodeToString(imageBytes, Base64.NO_WRAP);

                JSObject result = new JSObject();
                result.put("success", true);
                result.put("data", base64);

                new Handler(Looper.getMainLooper()).post(() -> call.resolve(result));

            } catch (Exception e) {
                Log.e(TAG, "LOKPlugin: getImageAsBase64 error", e);
                JSObject result = new JSObject();
                result.put("success", false);
                result.put("error", e.getMessage());
                new Handler(Looper.getMainLooper()).post(() -> call.resolve(result));
            }
        });
    }

    /**
     * Delete temporary image files
     */
    @PluginMethod()
    public void deleteImageFiles(PluginCall call) {
        try {
            org.json.JSONArray paths = call.getArray("paths");
            int deleted = 0;

            if (paths != null) {
                for (int i = 0; i < paths.length(); i++) {
                    String path = paths.getString(i);
                    File file = new File(path);
                    if (file.exists() && file.delete()) {
                        deleted++;
                    }
                }
            }

            JSObject result = new JSObject();
            result.put("success", true);
            result.put("deleted", deleted);
            call.resolve(result);
            Log.d(TAG, "LOKPlugin: Deleted " + deleted + " temp files");

        } catch (Exception e) {
            Log.e(TAG, "LOKPlugin: deleteImageFiles error", e);
            JSObject result = new JSObject();
            result.put("success", false);
            result.put("error", e.getMessage());
            call.resolve(result);
        }
    }
}
