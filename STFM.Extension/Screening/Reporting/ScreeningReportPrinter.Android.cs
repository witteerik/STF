// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

#if ANDROID
using Android.Content;
using Android.OS;
using Android.Print;

namespace STFM.Extension.Screening.Reporting;

public static partial class ScreeningReportPrinter
{
    private static readonly HashSet<ReportPrintAdapter> ActiveAdapters = [];

    private static Task<ScreeningPrintRequestResult> PrintAndroidAsync(ScreeningReportDocument document, CancellationToken cancellationToken)
        => MainThread.InvokeOnMainThreadAsync(async () =>
        {
            if (document.WritePages is null)
                throw new InvalidOperationException("Skapa rapporten igen innan den skrivs ut.");
            var activity = Platform.CurrentActivity
                ?? throw new InvalidOperationException("Ingen aktiv Android-vy finns för utskriftsdialogen.");
            var manager = activity.GetSystemService(Context.PrintService) as PrintManager
                ?? throw new InvalidOperationException("Androids utskriftstjänst är inte tillgänglig.");
            cancellationToken.ThrowIfCancellationRequested();
            var adapter = new ReportPrintAdapter(document, cancellationToken);
            lock (ActiveAdapters) ActiveAdapters.Add(adapter);
            try
            {
                var attributes = new PrintAttributes.Builder()
                    .SetMediaSize(PrintAttributes.MediaSize.IsoA4!)
                    .SetColorMode(PrintColorMode.Color)
                    .SetMinMargins(new PrintAttributes.Margins(0, 0, 0, 0))
                    .Build();
                var job = manager.Print("Hörselscreening", adapter, attributes)
                    ?? throw new InvalidOperationException("Utskriftsdialogen kunde inte öppnas.");
                // The camera/tests are already stopped by the host. Do not cancel
                // this wait on Window.Stopped: opening the print UI causes that event.
                await adapter.Finished;
                if (job.IsCancelled)
                    return new ScreeningPrintRequestResult(false, false, "Utskriften avbröts. Rapporten och patientuppgifterna finns kvar.");
                if (job.IsFailed || adapter.LastError is not null)
                    return new ScreeningPrintRequestResult(false, false, adapter.LastError ?? "Utskriftstjänsten rapporterade ett fel. Försök igen.");
                return new ScreeningPrintRequestResult(true, false,
                    "Utskriftsdialogen är stängd. Kontrollera att hela rapporten har kommit ur skrivaren innan du bekräftar och startar nästa patient.");
            }
            finally
            {
                lock (ActiveAdapters) ActiveAdapters.Remove(adapter);
                adapter.Dispose();
            }
        });

    private sealed class ReportPrintAdapter(ScreeningReportDocument document, CancellationToken requestedCancellation) : PrintDocumentAdapter
    {
        private readonly TaskCompletionSource finished = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private float pageWidth = ScreeningReportRenderer.A4Width;
        private float pageHeight = ScreeningReportRenderer.A4Height;
        private float margin;
        public Task Finished => finished.Task;
        public string? LastError { get; private set; }

        public override void OnLayout(PrintAttributes? oldAttributes, PrintAttributes? newAttributes,
            CancellationSignal? cancellationSignal, LayoutResultCallback? callback, Bundle? extras)
        {
            if (callback is null) return;
            if (cancellationSignal?.IsCanceled == true || requestedCancellation.IsCancellationRequested)
            {
                callback.OnLayoutCancelled();
                return;
            }
            try
            {
                var media = newAttributes?.GetMediaSize();
                if (media is not null)
                {
                    pageWidth = media.WidthMils * 72f / 1000;
                    pageHeight = media.HeightMils * 72f / 1000;
                }
                var margins = newAttributes?.MinMargins;
                margin = margins is null ? 0 : Math.Max(Math.Max(margins.LeftMils, margins.RightMils),
                    Math.Max(margins.TopMils, margins.BottomMils)) * 72f / 1000;
                if (pageWidth <= 2 * margin || pageHeight <= 2 * margin)
                    throw new InvalidOperationException("Skrivarens utskriftsyta är för liten.");
                using var info = new PrintDocumentInfo.Builder("hearing-screening.pdf")
                    .SetContentType(PrintContentType.Document)
                    .SetPageCount(document.PageCount)
                    .Build();
                callback.OnLayoutFinished(info, true);
                LastError = null;
            }
            catch (Exception error)
            {
                LastError = error.Message;
                callback.OnLayoutFailed(error.Message);
            }
        }

        public override async void OnWrite(PageRange[]? pages, ParcelFileDescriptor? destination,
            CancellationSignal? cancellationSignal, WriteResultCallback? callback)
        {
            if (callback is null) return;
            using var cancellation = CancellationTokenSource.CreateLinkedTokenSource(requestedCancellation);
            using var listener = new CancelListener(cancellation);
            cancellationSignal?.SetOnCancelListener(listener);
            try
            {
                if (cancellationSignal?.IsCanceled == true) cancellation.Cancel();
                cancellation.Token.ThrowIfCancellationRequested();
                if (destination is null) throw new IOException("Utskriftsmålet saknas.");
                var selected = Enumerable.Range(0, document.PageCount)
                    .Where(page => pages is not null && pages.Any(range => range.Start <= page && page <= range.End)).ToArray();
                if (selected.Length == 0) throw new InvalidOperationException("Inga rapportsidor valdes.");
                using var pdf = new MemoryStream();
                await Task.Run(() => document.WritePages!(pdf, selected, pageWidth, pageHeight, margin, cancellation.Token), cancellation.Token);
                cancellation.Token.ThrowIfCancellationRequested();
                // AutoCloseOutputStream owns and closes the framework's descriptor.
                using (var output = new ParcelFileDescriptor.AutoCloseOutputStream(destination))
                {
                    var bytes = pdf.GetBuffer();
                    for (int offset = 0; offset < pdf.Length; offset += 65536)
                    {
                        cancellation.Token.ThrowIfCancellationRequested();
                        int length = (int)Math.Min(65536, pdf.Length - offset);
                        output.Write(bytes, offset, length);
                    }
                    output.Flush();
                }
                cancellation.Token.ThrowIfCancellationRequested();
                callback.OnWriteFinished(selected.Select(page => new PageRange(page, page)).ToArray());
                LastError = null;
            }
            catch (System.OperationCanceledException)
            {
                TryClose(destination);
                callback.OnWriteCancelled();
            }
            catch (Exception error)
            {
                TryClose(destination);
                LastError = error.Message;
                callback.OnWriteFailed(error.Message);
            }
            finally
            {
                cancellationSignal?.SetOnCancelListener(null);
            }
        }

        public override void OnFinish()
        {
            finished.TrySetResult();
            base.OnFinish();
        }

        private static void TryClose(ParcelFileDescriptor? descriptor)
        {
            try { descriptor?.Close(); }
            catch (Exception) { /* The stream may already have closed the descriptor. */ }
        }

        private sealed class CancelListener(CancellationTokenSource cancellation) : Java.Lang.Object, CancellationSignal.IOnCancelListener
        {
            public void OnCancel()
            {
                try { cancellation.Cancel(); }
                catch (ObjectDisposedException) { }
            }
        }
    }
}
#endif
