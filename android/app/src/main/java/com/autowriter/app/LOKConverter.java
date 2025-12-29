/**
 * ===========================================
 * LibreOfficeKit Converter with Font Setup
 * ===========================================
 * 
 * Extracts fonts from APK and initializes LibreOffice.
 * 
 * @author AutoWriter
 * @version 5.0.0
 */
package com.autowriter.app;

import android.app.Activity;
import android.content.Context;
import android.content.res.AssetManager;
import android.util.Log;

import org.libreoffice.kit.LibreOfficeKit;
import org.libreoffice.kit.Office;
import org.libreoffice.kit.Document;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.ByteBuffer;

/**
 * LOKConverter - Extracts fonts and initializes LibreOffice
 */
public class LOKConverter {

    private static final String TAG = "LOKConverter";

    private Context context;
    private Office office;
    private boolean isReady = false;
    private String lastError = null;

    /**
     * Calculate dynamic page limit based on available memory
     * Each page takes approximately 3-5MB of memory when rendered
     */
    private int calculateMaxPages() {
        Runtime runtime = Runtime.getRuntime();
        long maxMem = runtime.maxMemory(); // Max heap size
        long usedMem = runtime.totalMemory() - runtime.freeMemory();
        long availableMem = maxMem - usedMem;

        // Estimate ~4MB per page for rendering + base64 encoding
        // Use only 50% of available memory to leave room for other operations
        long safeMemory = (availableMem / 2) / (4 * 1024 * 1024);

        // Minimum 10 pages, maximum 150 pages
        int maxPages = Math.max(10, Math.min(150, (int) safeMemory));

        Log.d(TAG, "Dynamic page limit: " + maxPages + " pages (available: " + (availableMem / 1024 / 1024) + "MB)");
        return maxPages;
    }

    /**
     * Constructor - initializes LibreOfficeKit
     */
    public LOKConverter(Context context) {
        this.context = context;
        initialize();
    }

    /**
     * Initialize LibreOfficeKit with font extraction
     */
    private void initialize() {
        if (!(context instanceof Activity)) {
            lastError = "Context must be an Activity";
            Log.e(TAG, lastError);
            return;
        }

        Activity activity = (Activity) context;

        try {
            Log.d(TAG, "Starting LibreOffice bootstrap...");

            // Extract fonts and config from assets/unpack/
            extractUnpackAssets(activity);

            // Initialize LibreOfficeKit
            Log.d(TAG, "Calling LibreOfficeKit.init()...");
            LibreOfficeKit.init(activity);

            // Get office handle
            ByteBuffer lokHandle = LibreOfficeKit.getLibreOfficeKitHandle();
            if (lokHandle != null) {
                office = new Office(lokHandle);
                isReady = true;
                Log.d(TAG, "LibreOfficeKit initialized successfully!");
            } else {
                lastError = "Failed to get LibreOfficeKit handle";
                Log.e(TAG, lastError);
            }

        } catch (UnsatisfiedLinkError e) {
            lastError = "Native library error: " + e.getMessage();
            Log.e(TAG, lastError, e);
        } catch (Exception e) {
            lastError = "Init error: " + e.getMessage();
            Log.e(TAG, lastError, e);
        }
    }

    /**
     * Extract assets/unpack/ folder to app data directory
     * This includes fonts.conf, fonts, and UNO type libraries
     */
    private void extractUnpackAssets(Activity activity) {
        String dataDir = activity.getApplicationInfo().dataDir;
        AssetManager assetManager = activity.getAssets();

        // Check if already extracted (v3 = official LibreOffice assets)
        File marker = new File(dataDir, ".lok_assets_v3");
        if (marker.exists()) {
            Log.d(TAG, "Assets already extracted");
            return;
        }

        Log.d(TAG, "Extracting LibreOffice assets to " + dataDir);

        try {
            // Extract etc/fonts/fonts.conf
            extractAssetFolder(assetManager, "unpack/etc", dataDir + "/etc");

            // Extract user/fonts/*.ttf
            extractAssetFolder(assetManager, "unpack/user", dataDir + "/user");

            // Extract program/*.rdb (UNO type libraries)
            extractAssetFolder(assetManager, "unpack/program", dataDir + "/program");

            // Create marker file
            marker.createNewFile();

            Log.d(TAG, "Assets extraction complete");
        } catch (Exception e) {
            Log.e(TAG, "Failed to extract assets", e);
        }
    }

    /**
     * Recursively extract an asset folder to disk
     */
    private void extractAssetFolder(AssetManager assetManager, String assetPath, String destPath) {
        try {
            String[] files = assetManager.list(assetPath);
            if (files == null || files.length == 0) {
                // This is a file, copy it
                copyAssetFile(assetManager, assetPath, destPath);
                return;
            }

            // This is a directory
            File dir = new File(destPath);
            dir.mkdirs();

            for (String file : files) {
                String srcPath = assetPath + "/" + file;
                String dstPath = destPath + "/" + file;
                extractAssetFolder(assetManager, srcPath, dstPath);
            }
        } catch (Exception e) {
            Log.e(TAG, "Error extracting " + assetPath, e);
        }
    }

    /**
     * Copy a single asset file to disk
     */
    private void copyAssetFile(AssetManager assetManager, String assetPath, String destPath) {
        File destFile = new File(destPath);
        destFile.getParentFile().mkdirs();

        try (InputStream in = assetManager.open(assetPath);
                OutputStream out = new FileOutputStream(destFile)) {
            byte[] buffer = new byte[8192];
            int read;
            while ((read = in.read(buffer)) != -1) {
                out.write(buffer, 0, read);
            }
            Log.d(TAG, "Extracted: " + destPath);
        } catch (Exception e) {
            Log.e(TAG, "Error copying " + assetPath + ": " + e.getMessage());
        }
    }

    /**
     * Convert document from DOCX to PDF
     */
    public boolean convertToPdf(String inputPath, String outputPath) {
        if (!isReady || office == null) {
            Log.e(TAG, "LibreOfficeKit not ready: " + lastError);
            return false;
        }

        try {
            Log.d(TAG, "Loading document: " + inputPath);

            Document document = office.documentLoad("file://" + inputPath);
            if (document == null) {
                String error = office.getError();
                Log.e(TAG, "Failed to load document: " + error);
                return false;
            }

            Log.d(TAG, "Document loaded, saving as PDF...");

            document.initializeForRendering();
            document.saveAs("file://" + outputPath, "pdf", "");

            Log.d(TAG, "Document saved as PDF");

            File pdfFile = new File(outputPath);
            if (pdfFile.exists() && pdfFile.length() > 0) {
                Log.d(TAG, "PDF created: " + pdfFile.length() + " bytes");
                return true;
            } else {
                Log.e(TAG, "PDF file not created");
                return false;
            }

        } catch (Exception e) {
            Log.e(TAG, "Conversion failed", e);
            return false;
        }
    }

    // NOTE: Old paintTile-based convertToImage() removed - now using
    // convertToImageViaPdf()

    // NOTE: Old paintTile-based convertToImages() removed - now using
    // convertToImagesViaPdf()

    /**
     * Check if converter is ready
     */
    public boolean isReady() {
        return isReady;
    }

    /**
     * Get error message
     */
    public String getError() {
        return lastError;
    }

    /**
     * Cleanup
     */
    public void cleanup() {
        if (office != null) {
            try {
                office.destroy();
            } catch (Exception e) {
                Log.e(TAG, "Error destroying office", e);
            }
            office = null;
        }
        isReady = false;
    }

    /**
     * Convert DOCX to single PNG image via PDF (DOCX → PDF → PNG)
     * Uses Android PdfRenderer for accurate page rendering
     * 
     * @param inputPath    Path to the input DOCX file
     * @param outputPath   Path for the output PNG file
     * @param combinePages If true, combine all pages into one image
     * @return true if conversion successful
     */
    public boolean convertToImageViaPdf(String inputPath, String outputPath, boolean combinePages) throws Exception {
        if (!isReady || office == null) {
            Log.e(TAG, "LibreOfficeKit not ready: " + lastError);
            return false;
        }

        File tempPdf = null;
        boolean needsCleanup = false;

        try {
            // Check if input is already a PDF
            if (inputPath.toLowerCase().endsWith(".pdf")) {
                // Input is already PDF, use it directly
                tempPdf = new File(inputPath);
                needsCleanup = false;
                Log.d(TAG, "Input is already PDF, skipping conversion");
            } else {
                // Step 1: Convert DOCX to PDF
                tempPdf = new File(inputPath.replace(".docx", "_temp.pdf").replace(".doc", "_temp.pdf"));
                needsCleanup = true;
                if (!convertToPdf(inputPath, tempPdf.getAbsolutePath())) {
                    Log.e(TAG, "Failed to convert DOCX to PDF");
                    return false;
                }
            }

            // Step 2: Render PDF to PNG using PdfRenderer
            android.os.ParcelFileDescriptor pfd = android.os.ParcelFileDescriptor.open(
                    tempPdf, android.os.ParcelFileDescriptor.MODE_READ_ONLY);
            android.graphics.pdf.PdfRenderer renderer = new android.graphics.pdf.PdfRenderer(pfd);

            int pageCount = renderer.getPageCount();
            Log.d(TAG, "PDF has " + pageCount + " pages");

            // Dynamic page limit based on available memory
            final int maxPages = calculateMaxPages();
            if (pageCount > maxPages) {
                renderer.close();
                pfd.close();
                Log.e(TAG, "PDF has too many pages: " + pageCount + " (max: " + maxPages + ")");
                throw new Exception(
                        "الملف كبير جداً (" + pageCount + " صفحة). الحد الأقصى " + maxPages + " صفحة للذاكرة المتاحة");
            }

            if (pageCount == 0) {
                renderer.close();
                pfd.close();
                return false;
            }

            // Calculate DPI based on page count to avoid bitmap size limits
            // Android has max bitmap height around 30,000 pixels and max size ~100MP
            int dpi = 150;
            if (combinePages && pageCount > 20) {
                // For 20+ pages, reduce DPI to stay under limits
                // Estimate: A4 page at 150dpi = ~1750px height
                // Max 30000px / estimated_page_height = max pages at that DPI
                if (pageCount > 50) {
                    dpi = 72; // Very low DPI for 50+ pages
                } else if (pageCount > 30) {
                    dpi = 100; // Low DPI for 30-50 pages
                } else {
                    dpi = 120; // Medium DPI for 20-30 pages
                }
                Log.d(TAG, "Reducing DPI to " + dpi + " for " + pageCount + " pages to avoid bitmap limits");
            }

            java.util.ArrayList<android.graphics.Bitmap> pageBitmaps = new java.util.ArrayList<>();
            int totalHeight = 0;
            int maxWidth = 0;

            // Render each page
            for (int i = 0; i < pageCount; i++) {
                android.graphics.pdf.PdfRenderer.Page page = renderer.openPage(i);

                // Calculate pixel dimensions at target DPI
                int pixelWidth = (int) (page.getWidth() * dpi / 72.0);
                int pixelHeight = (int) (page.getHeight() * dpi / 72.0);

                android.graphics.Bitmap bitmap = android.graphics.Bitmap.createBitmap(
                        pixelWidth, pixelHeight, android.graphics.Bitmap.Config.ARGB_8888);
                bitmap.eraseColor(android.graphics.Color.WHITE);

                page.render(bitmap, null, null, android.graphics.pdf.PdfRenderer.Page.RENDER_MODE_FOR_DISPLAY);
                page.close();

                pageBitmaps.add(bitmap);
                totalHeight += pixelHeight;
                if (pixelWidth > maxWidth)
                    maxWidth = pixelWidth;
            }

            renderer.close();
            pfd.close();

            // Step 3: Combine or save single page
            android.graphics.Bitmap finalBitmap;
            if (combinePages && pageCount > 1) {
                // Combine all pages vertically
                finalBitmap = android.graphics.Bitmap.createBitmap(maxWidth, totalHeight,
                        android.graphics.Bitmap.Config.ARGB_8888);
                android.graphics.Canvas canvas = new android.graphics.Canvas(finalBitmap);
                canvas.drawColor(android.graphics.Color.WHITE);

                int yOffset = 0;
                for (android.graphics.Bitmap pageBitmap : pageBitmaps) {
                    canvas.drawBitmap(pageBitmap, 0, yOffset, null);
                    yOffset += pageBitmap.getHeight();
                    pageBitmap.recycle();
                }
            } else {
                // Just use first page (for single image)
                finalBitmap = pageBitmaps.get(0);
                for (int i = 1; i < pageBitmaps.size(); i++) {
                    pageBitmaps.get(i).recycle();
                }
            }

            // Save PNG
            File outputFile = new File(outputPath);
            try (java.io.FileOutputStream fos = new java.io.FileOutputStream(outputFile)) {
                finalBitmap.compress(android.graphics.Bitmap.CompressFormat.PNG, 100, fos);
                fos.flush();
            }
            finalBitmap.recycle();

            // Cleanup temp PDF only if we created it
            if (needsCleanup && tempPdf != null) {
                tempPdf.delete();
            }

            if (outputFile.exists() && outputFile.length() > 0) {
                Log.d(TAG, "PNG created via PDF: " + outputFile.length() + " bytes");
                return true;
            }
            return false;

        } catch (Exception e) {
            Log.e(TAG, "Image conversion via PDF failed", e);
            // Only delete tempPdf if we created it
            if (needsCleanup && tempPdf != null)
                tempPdf.delete();
            throw e; // Re-throw to preserve error message
        }
    }

    /**
     * Convert DOCX to multiple PNG images via PDF (one per page)
     * 
     * @param inputPath Path to the input DOCX file
     * @param outputDir Directory for output PNG files
     * @param baseName  Base name for output files
     * @return Array of created file paths, or null if failed
     */
    public String[] convertToImagesViaPdf(String inputPath, String outputDir, String baseName) throws Exception {
        if (!isReady || office == null) {
            Log.e(TAG, "LibreOfficeKit not ready: " + lastError);
            return null;
        }

        File tempPdf = null;
        boolean needsCleanup = false;
        java.util.ArrayList<String> outputFiles = new java.util.ArrayList<>();

        try {
            // Check if input is already a PDF
            if (inputPath.toLowerCase().endsWith(".pdf")) {
                // Input is already PDF, use it directly
                tempPdf = new File(inputPath);
                needsCleanup = false;
                Log.d(TAG, "Input is already PDF, skipping conversion");
            } else {
                // Step 1: Convert DOCX to PDF
                tempPdf = new File(inputPath.replace(".docx", "_temp.pdf")
                        .replace(".doc", "_temp.pdf")
                        .replace(".pptx", "_temp.pdf")
                        .replace(".ppt", "_temp.pdf")
                        .replace(".xlsx", "_temp.pdf")
                        .replace(".xls", "_temp.pdf"));
                needsCleanup = true;
                if (!convertToPdf(inputPath, tempPdf.getAbsolutePath())) {
                    Log.e(TAG, "Failed to convert DOCX to PDF");
                    return null;
                }
            }

            // Step 2: Render PDF pages using PdfRenderer
            android.os.ParcelFileDescriptor pfd = android.os.ParcelFileDescriptor.open(
                    tempPdf, android.os.ParcelFileDescriptor.MODE_READ_ONLY);
            android.graphics.pdf.PdfRenderer renderer = new android.graphics.pdf.PdfRenderer(pfd);

            int pageCount = renderer.getPageCount();
            Log.d(TAG, "PDF has " + pageCount + " pages");

            // Dynamic page limit based on available memory
            final int maxPages = calculateMaxPages();
            if (pageCount > maxPages) {
                renderer.close();
                pfd.close();
                Log.e(TAG, "PDF has too many pages: " + pageCount + " (max: " + maxPages + ")");
                throw new Exception(
                        "الملف كبير جداً (" + pageCount + " صفحة). الحد الأقصى " + maxPages + " صفحة للذاكرة المتاحة");
            }

            if (pageCount == 0) {
                renderer.close();
                pfd.close();
                return null;
            }

            int dpi = 150;

            // Render each page
            for (int i = 0; i < pageCount; i++) {
                android.graphics.pdf.PdfRenderer.Page page = renderer.openPage(i);

                int pixelWidth = (int) (page.getWidth() * dpi / 72.0);
                int pixelHeight = (int) (page.getHeight() * dpi / 72.0);

                android.graphics.Bitmap bitmap = android.graphics.Bitmap.createBitmap(
                        pixelWidth, pixelHeight, android.graphics.Bitmap.Config.ARGB_8888);
                bitmap.eraseColor(android.graphics.Color.WHITE);

                page.render(bitmap, null, null, android.graphics.pdf.PdfRenderer.Page.RENDER_MODE_FOR_DISPLAY);
                page.close();

                // Save this page
                String outputPath = outputDir + "/" + baseName + "_page" + (i + 1) + ".png";
                File outputFile = new File(outputPath);
                try (java.io.FileOutputStream fos = new java.io.FileOutputStream(outputFile)) {
                    bitmap.compress(android.graphics.Bitmap.CompressFormat.PNG, 100, fos);
                    fos.flush();
                }
                bitmap.recycle();

                if (outputFile.exists() && outputFile.length() > 0) {
                    outputFiles.add(outputPath);
                    Log.d(TAG, "Page " + (i + 1) + " saved: " + outputFile.length() + " bytes");
                }
            }

            renderer.close();
            pfd.close();

            // Only delete tempPdf if we created it during conversion
            if (needsCleanup && tempPdf != null) {
                tempPdf.delete();
            }

            if (outputFiles.size() > 0) {
                Log.d(TAG, "Created " + outputFiles.size() + " page images via PDF");
                return outputFiles.toArray(new String[0]);
            }
            return null;

        } catch (Exception e) {
            Log.e(TAG, "Multi-page image conversion via PDF failed", e);
            // Only delete tempPdf if we created it
            if (needsCleanup && tempPdf != null)
                tempPdf.delete();
            throw e; // Re-throw to preserve error message
        }
    }

    /**
     * Progress callback interface for page-by-page conversion
     */
    public interface ProgressCallback {
        void onProgress(int current, int total);
    }

    /**
     * Convert document to images with progress callback
     * Processes pages one at a time to minimize memory usage
     */
    public String[] convertToImagesViaPdfWithProgress(String inputPath, String outputDir,
            String baseName, ProgressCallback callback) throws Exception {
        if (!isReady || office == null) {
            Log.e(TAG, "LibreOfficeKit not ready: " + lastError);
            return null;
        }

        File tempPdf = null;
        boolean needsCleanup = false;
        java.util.ArrayList<String> outputFiles = new java.util.ArrayList<>();

        try {
            // Check if input is already a PDF
            if (inputPath.toLowerCase().endsWith(".pdf")) {
                tempPdf = new File(inputPath);
                needsCleanup = false;
                Log.d(TAG, "Input is already PDF, skipping conversion");
            } else {
                // Convert to PDF first
                tempPdf = new File(inputPath.replace(".docx", "_temp.pdf")
                        .replace(".doc", "_temp.pdf")
                        .replace(".pptx", "_temp.pdf")
                        .replace(".ppt", "_temp.pdf")
                        .replace(".xlsx", "_temp.pdf")
                        .replace(".xls", "_temp.pdf"));
                needsCleanup = true;
                if (!convertToPdf(inputPath, tempPdf.getAbsolutePath())) {
                    Log.e(TAG, "Failed to convert document to PDF");
                    return null;
                }
            }

            // Render PDF pages one by one
            android.os.ParcelFileDescriptor pfd = android.os.ParcelFileDescriptor.open(
                    tempPdf, android.os.ParcelFileDescriptor.MODE_READ_ONLY);
            android.graphics.pdf.PdfRenderer renderer = new android.graphics.pdf.PdfRenderer(pfd);

            int pageCount = renderer.getPageCount();
            Log.d(TAG, "PDF has " + pageCount + " pages, processing with progress...");

            // NOTE: No page limit here - this function processes pages ONE at a time
            // and recycles memory after each page, so it can handle any number of pages

            if (pageCount == 0) {
                renderer.close();
                pfd.close();
                return null;
            }

            int dpi = 150;

            // Process ONE page at a time to minimize memory
            for (int i = 0; i < pageCount; i++) {
                android.graphics.pdf.PdfRenderer.Page page = renderer.openPage(i);

                int pixelWidth = (int) (page.getWidth() * dpi / 72.0);
                int pixelHeight = (int) (page.getHeight() * dpi / 72.0);

                android.graphics.Bitmap bitmap = android.graphics.Bitmap.createBitmap(
                        pixelWidth, pixelHeight, android.graphics.Bitmap.Config.ARGB_8888);
                bitmap.eraseColor(android.graphics.Color.WHITE);

                page.render(bitmap, null, null, android.graphics.pdf.PdfRenderer.Page.RENDER_MODE_FOR_DISPLAY);
                page.close();

                // Save immediately
                String outputPath = outputDir + "/" + baseName + "_page" + (i + 1) + ".png";
                File outputFile = new File(outputPath);
                try (java.io.FileOutputStream fos = new java.io.FileOutputStream(outputFile)) {
                    bitmap.compress(android.graphics.Bitmap.CompressFormat.PNG, 100, fos);
                    fos.flush();
                }

                // Recycle bitmap immediately to free memory
                bitmap.recycle();
                bitmap = null;

                if (outputFile.exists() && outputFile.length() > 0) {
                    outputFiles.add(outputPath);
                    Log.d(TAG, "Page " + (i + 1) + "/" + pageCount + " saved");
                }

                // Report progress
                if (callback != null) {
                    callback.onProgress(i + 1, pageCount);
                }
            }

            renderer.close();
            pfd.close();

            // Cleanup temp PDF
            if (needsCleanup && tempPdf != null) {
                tempPdf.delete();
            }

            if (outputFiles.size() > 0) {
                Log.d(TAG, "Created " + outputFiles.size() + " page images with progress");
                return outputFiles.toArray(new String[0]);
            }
            return null;

        } catch (Exception e) {
            Log.e(TAG, "Image conversion with progress failed", e);
            if (needsCleanup && tempPdf != null)
                tempPdf.delete();
            throw e; // Re-throw to preserve error message
        }
    }
}
