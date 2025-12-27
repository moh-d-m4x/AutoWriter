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
            java.io.InputStream is = getContentResolver().openInputStream(imageUri);
            java.io.File cacheDir = getCacheDir();

            // Try to convert mime to extension
            String mime = getContentResolver().getType(imageUri);
            String ext = ".bin"; // Default

            if (mime != null) {
                if (mime.contains("image/png"))
                    ext = ".png";
                else if (mime.contains("image/jpeg"))
                    ext = ".jpg";
                else if (mime.contains("image/webp"))
                    ext = ".webp";
                else if (mime.contains("wordprocessingml") || mime.contains("docx"))
                    ext = ".docx";
                else if (mime.contains("msword") || mime.contains("doc"))
                    ext = ".doc";
                else if (mime.contains("pdf"))
                    ext = ".pdf";
                else if (mime.contains("text/plain"))
                    ext = ".txt";
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
