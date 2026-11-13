# VisioForge Android Java glue

These files are the small Android companion project supplied with the official
[VisioForge .NET SDK samples](https://github.com/visioforge/.Net-SDK-s-samples/tree/8dcca4b1fcf2a1f90b68e3d7ad8631f9d9b0d4ce/AndroidDependency).

Pinned upstream commit: `8dcca4b1fcf2a1f90b68e3d7ad8631f9d9b0d4ce`.

The Java code and `VisioForge.Core.Android.X10.csproj` match upstream; only trailing
whitespace has been removed.
The .NET Android build compiles these Java sources into the application. There
is no separate Android Studio, Gradle or NDK build. The native runtime comes
from the pinned `VisioForge.CrossPlatform.Core.Android` NuGet dependency.

The files sit outside `STFM.Extension` to keep the MAUI project's default source
globs from accidentally including another project's generated files.

## Licences

- `LICENSE` is the upstream sample repository's MIT licence.
- The copyright and licence headers in the Java files are preserved. In
  particular, `GstAhcCallback.java` and `GstAmcOnFrameAvailableListener.java`
  specify LGPL 2.1; `GstAhsCallback.java` specifies LGPL 2 or later. Copies of
  those licence texts are included here. These specific notices take precedence
  over the sample repository's general licence for those files.
- The VisioForge SDK itself is separately licensed. Refer to its trial or
  commercial licence and the third-party notices distributed in its NuGet
  packages. Copying this companion project does not grant a VisioForge SDK
  licence.
