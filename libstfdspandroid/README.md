# libstfdspandroid

Modern Android NDK + CMake replacement for the legacy `libostfdspandroid` project.

## Layout

- `CMakeLists.txt` - shared-library build definition for Android NDK builds
- `include/` - public native headers
- `src/` - native implementation files

## Compatibility

The CMake target is named `libstfdspandroid`, while the produced shared library keeps the compatibility output name `libostfdspandroid` so the existing Android P/Invoke usage in `STFN.Core` and packaging in `StfResearchSuite` can continue to work without changing the runtime native library lookup contract.

## Build requirements

- Android SDK installed
- Android NDK with CMake support installed
- CMake 3.22.1 or newer available through the Android SDK or system installation
- Ninja available on `PATH` for the configured CMake generator

The project is intended to be built with the Android NDK toolchain. It is not wired into the Visual Studio solution as an MSBuild C++ Android project.

## Automated build/publish script

Use `build-android-native.ps1` to build all required ABIs and publish the outputs into `StfResearchSuite/Platforms/Android/native-libs/`.

Default behavior:

- Resolves Android SDK from `-AndroidSdkRoot`, `ANDROID_SDK_ROOT`, `ANDROID_HOME`, or `%LOCALAPPDATA%\Android\Sdk`
- Resolves Android NDK from `-AndroidNdkRoot`, `ANDROID_NDK_ROOT`, `ANDROID_NDK_HOME`, `<sdk>\ndk-bundle`, or the newest directory under `<sdk>\ndk`
- Builds `armeabi-v7a`, `arm64-v8a`, `x86`, and `x86_64`
- Uses `Release` and `android-31`
- Copies each built `liblibostfdspandroid.so` into the matching MAUI Android native-libs ABI folder

Example usage:

- `powershell -ExecutionPolicy Bypass -File .\libstfdspandroid\build-android-native.ps1`
- `powershell -ExecutionPolicy Bypass -File .\libstfdspandroid\build-android-native.ps1 -Clean`
- `powershell -ExecutionPolicy Bypass -File .\libstfdspandroid\build-android-native.ps1 -AndroidSdkRoot C:\Android\Sdk -AndroidNdkRoot C:\Android\Sdk\ndk\28.0.13004108`

Useful options:

- `-Configuration Debug|Release`
- `-AndroidPlatform android-31`
- `-Abis armeabi-v7a,arm64-v8a`
- `-SkipPublish`

## Integration with StfResearchSuite

1. Build `libstfdspandroid` for each required ABI.
2. Copy each generated `liblibostfdspandroid.so` into the matching ABI folder under `StfResearchSuite/Platforms/Android/native-libs/`.
3. Build `StfResearchSuite` for Android.

`StfResearchSuite.csproj` now includes Android native libraries from `Platforms/Android/native-libs/**/liblibostfdspandroid.so` and declares the supported ABIs explicitly.

### Optional automatic invocation from the MAUI build

`StfResearchSuite.csproj` contains an opt-in target that runs the script before Android builds when `BuildAndroidNativeDsp=true`.

Example:

- `dotnet build .\StfResearchSuite\StfResearchSuite.csproj -f net10.0-android -p:BuildAndroidNativeDsp=true`

The default is `false` so normal builds do not require the Android NDK to be installed.

## Assumptions

- Existing managed interop in `STFN.Core/libostfdsp_VB.vb` should remain unchanged for now.
- The authoritative Android FFT signature is the legacy Android variant that takes precomputed cosine and sine lookup arrays.
- The checked-in legacy `libostfdspandroid` source folder is retained only as a reference during migration.

## Notes

- The deprecated Visual C++ Android project file has been removed from active use.
- The new native code is organized under `include/` and `src/` for a cleaner NDK/CMake layout.
- If you later want the managed layer to use the new project identity directly, update Android `DllImport` names and change the CMake output name at the same time.
