namespace DhcbTools.Core.AutoCAD.DrawingCleanup;

/// <summary>Cấu hình lưu đè bản vẽ đang mở về đúng file của nó (như QSAVE).</summary>
public sealed class DrawingSaveConfig
{
    /// <summary>Chỉ xem trước: báo sẽ lưu file nào, phiên bản DWG nào — không ghi.</summary>
    public bool DryRun { get; init; } = true;

    /// <summary>Giữ bản trước khi lưu thành file .bak cạnh bản vẽ, như lưu tay (mặc định bật).</summary>
    public bool CreateBackup { get; init; } = true;
}
