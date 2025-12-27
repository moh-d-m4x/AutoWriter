// Icon generation script - generates all required icon sizes from SVG
// Uses trim to auto-detect content bounds, then centers on canvas
const sharp = require('sharp');
const fs = require('fs');
const path = require('path');
const pngToIco = require('png-to-ico').default || require('png-to-ico');

const svgPath = path.join(__dirname, 'autowriter_icon.svg');
let svgContent = fs.readFileSync(svgPath, 'utf8');

// Remove dark reader attributes that interfere with rendering
svgContent = svgContent.replace(/data-darkreader-[^=]+="[^"]*"/g, '');
svgContent = svgContent.replace(/--darkreader-[^;]+;/g, '');

const svgBuffer = Buffer.from(svgContent);

// All icon configurations
const icons = [
    // Electron icons
    { path: 'electron/icon.png', size: 512 },
    { path: 'electron/icon.ico', size: 256, format: 'ico' },

    // Android launcher icons
    { path: 'android/app/src/main/res/mipmap-mdpi/ic_launcher.png', size: 48 },
    { path: 'android/app/src/main/res/mipmap-mdpi/ic_launcher_round.png', size: 48 },
    { path: 'android/app/src/main/res/mipmap-hdpi/ic_launcher.png', size: 72 },
    { path: 'android/app/src/main/res/mipmap-hdpi/ic_launcher_round.png', size: 72 },
    { path: 'android/app/src/main/res/mipmap-xhdpi/ic_launcher.png', size: 96 },
    { path: 'android/app/src/main/res/mipmap-xhdpi/ic_launcher_round.png', size: 96 },
    { path: 'android/app/src/main/res/mipmap-xxhdpi/ic_launcher.png', size: 144 },
    { path: 'android/app/src/main/res/mipmap-xxhdpi/ic_launcher_round.png', size: 144 },
    { path: 'android/app/src/main/res/mipmap-xxxhdpi/ic_launcher.png', size: 192 },
    { path: 'android/app/src/main/res/mipmap-xxxhdpi/ic_launcher_round.png', size: 192 },

    // Android adaptive icon foregrounds (with padding for safe zone)
    { path: 'android/app/src/main/res/mipmap-mdpi/ic_launcher_foreground.png', size: 108, iconSize: 72 },
    { path: 'android/app/src/main/res/mipmap-hdpi/ic_launcher_foreground.png', size: 162, iconSize: 108 },
    { path: 'android/app/src/main/res/mipmap-xhdpi/ic_launcher_foreground.png', size: 216, iconSize: 144 },
    { path: 'android/app/src/main/res/mipmap-xxhdpi/ic_launcher_foreground.png', size: 324, iconSize: 216 },
    { path: 'android/app/src/main/res/mipmap-xxxhdpi/ic_launcher_foreground.png', size: 432, iconSize: 288 },

    // Splash screen images - larger sizes for better visibility
    { path: 'android/app/src/main/res/drawable/splash.png', size: 384 },
    { path: 'android/app/src/main/res/drawable-port-mdpi/splash.png', size: 288 },
    { path: 'android/app/src/main/res/drawable-port-hdpi/splash.png', size: 384 },
    { path: 'android/app/src/main/res/drawable-port-xhdpi/splash.png', size: 512 },
    { path: 'android/app/src/main/res/drawable-port-xxhdpi/splash.png', size: 640 },
    { path: 'android/app/src/main/res/drawable-port-xxxhdpi/splash.png', size: 768 },
    { path: 'android/app/src/main/res/drawable-land-mdpi/splash.png', size: 288 },
    { path: 'android/app/src/main/res/drawable-land-hdpi/splash.png', size: 384 },
    { path: 'android/app/src/main/res/drawable-land-xhdpi/splash.png', size: 512 },
    { path: 'android/app/src/main/res/drawable-land-xxhdpi/splash.png', size: 640 },
    { path: 'android/app/src/main/res/drawable-land-xxxhdpi/splash.png', size: 768 },
];

async function generateIcon(config, baseImage) {
    const fullPath = path.join(__dirname, config.path);
    const dir = path.dirname(fullPath);

    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }

    try {
        // Handle ICO format separately
        if (config.format === 'ico') {
            // Generate multiple sizes for ICO (256, 128, 64, 48, 32, 16)
            const sizes = [256, 128, 64, 48, 32, 16];
            const pngBuffers = [];

            for (const size of sizes) {
                const resizedBuffer = await sharp(baseImage)
                    .resize(size, size, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
                    .png()
                    .toBuffer();
                pngBuffers.push(resizedBuffer);
            }

            const icoBuffer = await pngToIco(pngBuffers);
            fs.writeFileSync(fullPath, icoBuffer);
            console.log(`✓ ${config.path} (ICO with ${sizes.join(', ')}px)`);
            return;
        }

        if (config.iconSize) {
            // Adaptive foreground: render icon smaller and center on transparent canvas
            const iconBuffer = await sharp(baseImage)
                .resize(config.iconSize, config.iconSize, { fit: 'inside' })
                .png()
                .toBuffer();

            // Get actual dimensions after resize
            const meta = await sharp(iconBuffer).metadata();
            const topPad = Math.floor((config.size - meta.height) / 2);
            const leftPad = Math.floor((config.size - meta.width) / 2);

            await sharp({
                create: {
                    width: config.size,
                    height: config.size,
                    channels: 4,
                    background: { r: 0, g: 0, b: 0, alpha: 0 }
                }
            })
                .composite([{ input: iconBuffer, left: leftPad, top: topPad }])
                .png()
                .toFile(fullPath);

            console.log(`✓ ${config.path} (${config.size}x${config.size}, icon: ${config.iconSize}px)`);
        } else {
            // Regular icon: resize to fit inside, then center
            const resizedBuffer = await sharp(baseImage)
                .resize(config.size, config.size, { fit: 'inside' })
                .png()
                .toBuffer();

            // Get actual dimensions
            const meta = await sharp(resizedBuffer).metadata();

            if (meta.width === config.size && meta.height === config.size) {
                // Already correct size
                await sharp(resizedBuffer).toFile(fullPath);
            } else {
                // Need to center on canvas
                const topPad = Math.floor((config.size - meta.height) / 2);
                const leftPad = Math.floor((config.size - meta.width) / 2);

                await sharp({
                    create: {
                        width: config.size,
                        height: config.size,
                        channels: 4,
                        background: { r: 0, g: 0, b: 0, alpha: 0 }
                    }
                })
                    .composite([{ input: resizedBuffer, left: leftPad, top: topPad }])
                    .png()
                    .toFile(fullPath);
            }

            console.log(`✓ ${config.path} (${config.size}x${config.size})`);
        }
    } catch (err) {
        console.error(`✗ ${config.path}: ${err.message}`);
    }
}

async function main() {
    console.log('Generating icons from autowriter_icon.svg...\n');

    // First, render SVG to a large PNG, trim whitespace, and this becomes our base
    console.log('Step 1: Rendering SVG at high resolution...');

    // Render at 1024px first
    let baseImage = await sharp(svgBuffer, { density: 300 })
        .resize(1024, 1024)
        .png()
        .toBuffer();

    // Trim any transparent edge to get actual content bounds
    console.log('Step 2: Trimming to content bounds...');
    const trimmed = await sharp(baseImage)
        .trim()
        .png()
        .toBuffer();

    // Get trimmed dimensions
    const trimMeta = await sharp(trimmed).metadata();
    console.log(`   Trimmed size: ${trimMeta.width}x${trimMeta.height}`);

    // Now extend back to square, centered, with 20% extra canvas
    console.log('Step 3: Centering on square canvas (with 20% extra padding)...');
    const maxDim = Math.max(trimMeta.width, trimMeta.height);
    const canvasSize = Math.round(maxDim * 1); // 20% larger canvas (edited to 0% for now)
    const topPad = Math.floor((canvasSize - trimMeta.height) / 2);
    const leftPad = Math.floor((canvasSize - trimMeta.width) / 2);

    baseImage = await sharp({
        create: {
            width: canvasSize,
            height: canvasSize,
            channels: 4,
            background: { r: 0, g: 0, b: 0, alpha: 0 }
        }
    })
        .composite([{ input: trimmed, left: leftPad, top: topPad }])
        .png()
        .toBuffer();

    console.log(`   Base image: ${canvasSize}x${canvasSize} (icon content: ${maxDim}px)\n`);

    // Now generate all icons from this centered base
    console.log('Step 4: Generating icons...');
    for (const config of icons) {
        await generateIcon(config, baseImage);
    }

    console.log('\nDone! All icons generated.');
}

main();
