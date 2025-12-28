package com.autowriter.app;

import com.getcapacitor.BridgeActivity;
import android.os.Bundle;
import androidx.core.splashscreen.SplashScreen;

/**
 * LibreOfficeKit code implementation
 * MainActivity registers the LOKPlugin for document conversion
 */
public class MainActivity extends BridgeActivity {

    // Track processed URIs to prevent duplicate processing
    private static final java.util.Set<String> processedIntentUris = java.util.Collections
            .synchronizedSet(new java.util.HashSet<>());

    @Override
    public void onCreate(Bundle savedInstanceState) {
        // Install splash screen before calling super.onCreate()
        SplashScreen.installSplashScreen(this);

        // LibreOfficeKit code implementation - Register LOK plugin
        registerPlugin(LOKPlugin.class);
        super.onCreate(savedInstanceState);

        // Handle share intent if app launched via share
        handleIntent(getIntent());
    }

    @Override
    protected void onNewIntent(android.content.Intent intent) {
        super.onNewIntent(intent);
        handleIntent(intent);
    }

    private void handleIntent(android.content.Intent intent) {
        if (intent == null)
            return;

        String action = intent.getAction();
        String type = intent.getType();

        // Create a unique key for this intent based on action, data, and type
        String intentKey = (action != null ? action : "") + "|" +
                (intent.getData() != null ? intent.getData().toString() : "") + "|" +
                (type != null ? type : "");

        // Skip if we already fully processed this exact intent
        if (processedIntentUris.contains(intentKey)) {
            android.util.Log.d("MainActivity", "Skipping already-processed intent: " + intentKey);
            return;
        }

        if (android.content.Intent.ACTION_SEND.equals(action) && type != null) {
            android.net.Uri imageUri = (android.net.Uri) intent
                    .getParcelableExtra(android.content.Intent.EXTRA_STREAM);
            if (imageUri != null) {
                processSharedUri(imageUri);
                // Consume the extra to prevent duplicate processing
                intent.removeExtra(android.content.Intent.EXTRA_STREAM);
                // Mark this intent as fully processed
                processedIntentUris.add(intentKey);
            }
        } else if (android.content.Intent.ACTION_SEND_MULTIPLE.equals(action) && type != null) {
            java.util.ArrayList<android.net.Uri> imageUris = intent
                    .getParcelableArrayListExtra(android.content.Intent.EXTRA_STREAM);
            if (imageUris != null) {
                for (android.net.Uri uri : imageUris) {
                    processSharedUri(uri);
                }
                // Consume the extra to prevent duplicate processing
                intent.removeExtra(android.content.Intent.EXTRA_STREAM);
                // Mark this intent as fully processed
                processedIntentUris.add(intentKey);
            }
        } else if (android.content.Intent.ACTION_VIEW.equals(action)) {
            android.net.Uri data = intent.getData();
            if (data != null) {
                processSharedUri(data);
                // Consume the data to prevent duplicate processing
                intent.setData(null);
                // Mark this intent as fully processed
                processedIntentUris.add(intentKey);
            }
        }

        // Trigger event if we have pending images
        if (LOKPlugin.pendingSharedImagePaths != null && !LOKPlugin.pendingSharedImagePaths.isEmpty()) {
            if (getBridge() != null) {
                getBridge().triggerWindowJSEvent("appSharedImageAvailable", "{}");
            }
        }
    }

    private void processSharedUri(android.net.Uri imageUri) {
        // Check if this URI was already processed
        String uriString = imageUri.toString();
        if (processedIntentUris.contains(uriString)) {
            android.util.Log.d("MainActivity", "Skipping duplicate URI: " + uriString);
            return;
        }
        processedIntentUris.add(uriString);

        try {
            // Get MIME type first to filter unsupported files BEFORE loading
            String mime = getContentResolver().getType(imageUri);

            // Supported MIME types - skip loading unsupported files to prevent OOM
            java.util.Set<String> supportedMimes = new java.util.HashSet<>(java.util.Arrays.asList(
                    "image/png", "image/jpeg", "image/jpg", "image/webp", "image/gif", "image/bmp",
                    "application/pdf",
                    "application/msword", "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "application/vnd.ms-powerpoint",
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    "application/vnd.ms-excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

            boolean isSupported = false;
            if (mime != null) {
                // Check if MIME type starts with "image/" or matches exact supported types
                if (mime.startsWith("image/") || supportedMimes.contains(mime)) {
                    isSupported = true;
                }
            }

            if (!isSupported) {
                android.util.Log.d("MainActivity", "Skipping unsupported file type: " + mime);
                return; // Skip without loading to prevent OOM
            }

            java.io.InputStream is = getContentResolver().openInputStream(imageUri);
            java.io.File cacheDir = getCacheDir();

            // Convert MIME to extension
            String ext = ".bin"; // Default
            if (mime != null) {
                if (mime.contains("image/png"))
                    ext = ".png";
                else if (mime.contains("image/jpeg") || mime.contains("image/jpg"))
                    ext = ".jpg";
                else if (mime.contains("image/webp"))
                    ext = ".webp";
                else if (mime.contains("image/gif"))
                    ext = ".gif";
                else if (mime.contains("image/bmp"))
                    ext = ".bmp";
                else if (mime.contains("wordprocessingml") || mime.contains("docx"))
                    ext = ".docx";
                else if (mime.contains("msword") || mime.contains("doc"))
                    ext = ".doc";
                else if (mime.contains("presentationml") || mime.contains("pptx"))
                    ext = ".pptx";
                else if (mime.contains("ms-powerpoint") || mime.contains("ppt"))
                    ext = ".ppt";
                else if (mime.contains("spreadsheetml") || mime.contains("xlsx"))
                    ext = ".xlsx";
                else if (mime.contains("ms-excel") || mime.contains("xls"))
                    ext = ".xls";
                else if (mime.contains("pdf"))
                    ext = ".pdf";
            }

            // Generate filename with correct extension
            String fileName = "shared_file_" + System.currentTimeMillis() + "_" + (int) (Math.random() * 1000) + ext;

            java.io.File tempFile = new java.io.File(cacheDir, fileName);

            java.io.FileOutputStream fos = new java.io.FileOutputStream(tempFile);
            byte[] buffer = new byte[8192];
            int length;
            while ((length = is.read(buffer)) > 0) {
                fos.write(buffer, 0, length);
            }
            fos.close();
            is.close();

            // Add path to LOKPlugin list
            LOKPlugin.pendingSharedImagePaths.add(tempFile.getAbsolutePath());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
