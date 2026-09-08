# Photo Excel Sheet - Android build project

This package contains the working web app and Capacitor configuration.

## Easiest build method: Android Studio

1. Install Node.js LTS.
2. Install Android Studio from https://developer.android.com/studio
3. Open a terminal in this folder.
4. Run:
   npm install
   npx cap add android
   npx cap sync android
   npx cap open android
5. In Android Studio wait for Gradle sync.
6. Select **Build > Build APK(s)**.
7. The APK will be under:
   `android/app/build/outputs/apk/debug/app-debug.apk`

For a release APK, use Android Studio's **Build > Generate Signed Bundle / APK**.

## Important
The web app currently loads ExcelJS from jsDelivr. Therefore the app needs internet access for the Excel library unless ExcelJS is bundled locally later.

## GitHub Actions
A workflow is included in `.github/workflows/build-apk.yml`. After uploading this project to GitHub, the workflow can build a debug APK and publish it as a workflow artifact.
