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

                    for (String path : pendingSharedImagePaths) {
                        File imageFile = new File(path);
                        if (imageFile.exists()) {
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
}
