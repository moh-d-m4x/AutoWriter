# LibreOfficeKit Native Libraries Setup

## ⚠️ Important Note

The `.so` files in this directory were **built from LibreOffice source code** - NOT extracted from an APK. Extracting from APK does not work properly.

## Required Files

Place the native libraries here:

```
jniLibs/
├── arm64-v8a/
│   ├── liblo-native-code.so      (~182 MB)
│   ├── libc++_shared.so
│   ├── libfreebl3.so
│   ├── libnspr4.so
│   ├── libnss3.so
│   ├── libnssckbi.so
│   ├── libnssdbm3.so
│   ├── libnssutil3.so
│   ├── libplc4.so
│   ├── libplds4.so
│   ├── libsmime3.so
│   ├── libsoftokn3.so
│   ├── libsqlite3.so
│   └── libssl3.so
└── armeabi-v7a/
    └── (same files for 32-bit devices)
```

## How to Build from Source

Building LibreOfficeKit from source takes approximately **1.5 hours**.

### Prerequisites (Linux/WSL recommended)

1. **Install build dependencies:**
   ```bash
   sudo apt-get install git build-essential zip ccache \
       autoconf automake libtool pkg-config \
       openjdk-17-jdk ant
   ```

2. **Install Android SDK and NDK:**
   - Android SDK (API level 21+)
   - Android NDK r25 or compatible version

### Build Steps

1. **Clone LibreOffice core:**
   ```bash
   git clone https://gerrit.libreoffice.org/core libreoffice-core
   cd libreoffice-core
   ```

2. **Configure for Android:**
   ```bash
   ./autogen.sh --enable-release-build \
       --with-android-ndk=/path/to/ndk \
       --with-android-sdk=/path/to/sdk \
       --with-distro=LibreOfficeAndroid
   ```

3. **Build:**
   ```bash
   make
   ```

4. **Find the output:**
   The `.so` files will be in:
   ```
   instdir/program/
   ```

### Official Documentation

- [LibreOffice Android Build Guide](https://wiki.documentfoundation.org/Development/Android)
- [LibreOffice Core Repository](https://gerrit.libreoffice.org/plugins/gitiles/core)

## Pre-built Libraries Included

The `.so` files in this directory are **already built from the LibreOfficeKit core source code**.
You can use them directly - no need to build from source yourself!

## Note

These libraries total ~200MB. For most modern phones, only `arm64-v8a` is needed.
The build process requires a Linux environment (WSL works on Windows).
