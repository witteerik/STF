// Copyright (c) 2026 Erik Witte
// See License.txt in the repository root for this project's licence.

using Microsoft.Maui.Hosting;
#if ANDROID
using SkiaSharp.Views.Maui.Controls.Hosting;
using VisioForge.Core.UI.MAUI;
#endif

namespace STFM.Extension;

public static class MauiAppBuilderExtensions
{
    /// <summary>
    /// Registers the Android video preview used by <see cref="Views.OtoscopyView"/>.
    /// Call once while constructing the host application. The camera and SDK are
    /// opened lazily by the view, so startup does not request camera permission.
    /// </summary>
    public static MauiAppBuilder UseOtoscopy(this MauiAppBuilder builder)
    {
#if ANDROID
        builder.UseSkiaSharp()
            .ConfigureMauiHandlers(handlers => handlers.AddVisioForgeHandlers());
#endif
        return builder;
    }
}
