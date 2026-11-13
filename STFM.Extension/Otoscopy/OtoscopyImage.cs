// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

namespace STFM.Extension.Otoscopy;

public enum OtoscopyEar
{
    Left,
    Right
}

/// <summary>An otoscopy snapshot in JPEG format, retained in memory until the host saves it.</summary>
public sealed class OtoscopyImage
{
    private readonly byte[] jpegData;

    internal OtoscopyImage(OtoscopyEar ear, byte[] jpegData, DateTimeOffset capturedAt)
    {
        Ear = ear;
        this.jpegData = (byte[])jpegData.Clone();
        CapturedAt = capturedAt;
    }

    public OtoscopyEar Ear { get; }
    public DateTimeOffset CapturedAt { get; }
    public ReadOnlyMemory<byte> JpegData => jpegData;

    /// <summary>Returns a new read-only stream. The caller owns the stream.</summary>
    public Stream OpenRead() => new MemoryStream(jpegData, writable: false);
}

public sealed class OtoscopyImageCapturedEventArgs(OtoscopyImage image) : EventArgs
{
    public OtoscopyImage Image { get; } = image;
}
