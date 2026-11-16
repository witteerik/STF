using CommunityToolkit.Maui.Core.Primitives;
using CommunityToolkit.Maui.Storage;
using Microsoft.Maui.ApplicationModel;

namespace STFM.Storage;

internal static class StfFolderPicker
{
    public static async Task<FolderPickerResult> PickAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            return await MainThread.InvokeOnMainThreadAsync(async () =>
            {
                cancellationToken.ThrowIfCancellationRequested();
#if WINDOWS
                // Toolkit 15.0.1 uses Process.MainWindowHandle on .NET 10. Use
                // STF's actual MAUI window and the Windows App SDK picker instead.
                // This also supports running the Windows app as administrator.
                var window = Microsoft.Maui.Controls.Application.Current?.Windows
                    .Select(w => w.Handler?.PlatformView)
                    .OfType<Microsoft.UI.Xaml.Window>()
                    .FirstOrDefault();

                if (window is null)
                {
                    throw new InvalidOperationException(
                        "Cannot open the folder picker because the MAUI window is not ready.");
                }

                var picker = new Microsoft.Windows.Storage.Pickers.FolderPicker(window.AppWindow.Id)
                {
                    SuggestedStartLocation = Microsoft.Windows.Storage.Pickers.PickerLocationId.DocumentsLibrary
                };

                var operation = picker.PickSingleFolderAsync();
                await using var registration = cancellationToken.Register(() => operation.Cancel());
                var folder = await operation;
                cancellationToken.ThrowIfCancellationRequested();

                if (folder is null)
                {
                    throw new OperationCanceledException("Folder selection was cancelled.");
                }

                if (string.IsNullOrWhiteSpace(folder.Path))
                {
                    throw new IOException("The selected folder did not provide a file-system path.");
                }

                return new FolderPickerResult(new Folder(folder.Path, new DirectoryInfo(folder.Path).Name), null);
#else
                return await FolderPicker.Default.PickAsync(cancellationToken);
#endif
            });
        }
        catch (Exception exception)
        {
            return new FolderPickerResult(null, exception);
        }
    }
}
