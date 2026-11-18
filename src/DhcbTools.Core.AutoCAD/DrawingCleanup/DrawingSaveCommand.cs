using Autodesk.AutoCAD.DatabaseServices;
using DhcbTools.Shared.Logic.Cad;

namespace DhcbTools.Core.AutoCAD.DrawingCleanup;

/// <summary>
/// Lưu bản vẽ đang mở đè lên chính file của nó (như QSAVE) — để agent sửa qua Bridge xong tự lưu, không phải nhờ kỹ sư
/// bấm Ctrl+S hay gửi phím vào cửa sổ AutoCAD.
/// <para>
/// Giữ phiên bản DWG gốc (<see cref="Database.OriginalFileVersion"/>) để không âm thầm nâng định dạng file người khác
/// còn mở bằng AutoCAD cũ hơn, và để AutoCAD tạo .bak như lưu tay. Chỉ lưu tại chỗ: không nhận đường dẫn khác
/// (xem <see cref="DrawingSavePlanner"/>). Bản vẽ mới chưa đặt tên, file chỉ đọc hoặc file đã biến mất thì từ chối.
/// </para>
/// </summary>
public sealed class DrawingSaveCommand : ICoreCommand<DrawingSaveConfig>
{
    public string CommandName => "DrawingSave";

    public CommandResult Execute(Database database, DrawingSaveConfig config)
    {
        var fileName = database.Filename;
        var exists = !string.IsNullOrWhiteSpace(fileName) && File.Exists(fileName);
        var readOnly = exists && new FileInfo(fileName).IsReadOnly;
        if (DrawingSavePlanner.Problem(fileName, exists, readOnly) is { } problem)
        {
            return CommandResult.Fail(problem);
        }

        var version = database.OriginalFileVersion;
        if (config.DryRun)
        {
            return CommandResult.Ok($"[Xem trước] Sẽ lưu đè \"{fileName}\" (giữ định dạng {version}{(config.CreateBackup ? ", tạo .bak" : ", không tạo .bak")}). Chưa ghi.", 1);
        }

        var before = File.GetLastWriteTimeUtc(fileName);
        try
        {
            database.SaveAs(fileName, config.CreateBackup, version, database.SecurityParameters);
        }
        catch (Autodesk.AutoCAD.Runtime.Exception ex)
        {
            return CommandResult.Fail($"Không lưu được \"{fileName}\" ({ex.ErrorStatus}) — file có thể đang bị chương trình khác khoá.");
        }

        var after = File.GetLastWriteTimeUtc(fileName);
        var size = new FileInfo(fileName).Length;
        var result = CommandResult.Ok($"Đã lưu \"{fileName}\" ({size / 1024:N0} KB, định dạng {version}).", 1);
        if (after <= before)
        {
            // SaveAs không ném lỗi mà file không đổi giờ ghi thì không dám báo "đã lưu" suông.
            return CommandResult.Fail($"AutoCAD báo lưu xong nhưng \"{fileName}\" không đổi giờ ghi ({after:u}) — kiểm tra lại bằng tay.");
        }

        return result;
    }
}
