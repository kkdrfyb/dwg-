# Office Toolbox

Flutter rebuild of the Office Toolbox app.

## Quick Start

```bash
flutter pub get
flutter run
```

## Android Release Build

1. Update `namespace` and `applicationId` in `android/app/build.gradle.kts` to your own unique package name.
2. Create a keystore and signing config:

```bash
# Run from the android directory
mkdir keystore
keytool -genkeypair -v -keystore keystore/office_toolbox.jks -keyalg RSA -keysize 2048 -validity 10000 -alias office_toolbox
```

3. Copy `android/key.properties.example` to `android/key.properties` and fill in:

```properties
storeFile=../keystore/office_toolbox.jks
storePassword=YOUR_STORE_PASSWORD
keyAlias=office_toolbox
keyPassword=YOUR_KEY_PASSWORD
```

4. Build release artifacts:

```bash
flutter build apk --release
# or
flutter build appbundle --release
```

Notes:
- If `android/key.properties` is missing, the build falls back to debug signing for local testing.
- For production distribution, always use a real release keystore.
