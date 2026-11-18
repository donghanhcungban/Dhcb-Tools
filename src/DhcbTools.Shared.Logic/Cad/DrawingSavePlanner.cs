using System;
using System.IO;

namespace DhcbTools.Shared.Logic.Cad
{
    /// <summary>
    /// Kiểm điều kiện cho lệnh <c>DrawingSave</c> (lưu đè bản vẽ đang mở về đúng file của nó) — thuần, có test.
    /// Chỉ lưu tại chỗ: không nhận đường dẫn khác, để agent không thể dùng lệnh này ghi đè một file bất kỳ.
    /// </summary>
    public static class DrawingSavePlanner
    {
        /// <summary>
        /// Lý do không lưu được, hoặc null nếu lưu được. <paramref name="fileName"/> là <c>Database.Filename</c>:
        /// bản vẽ mới chưa đặt tên có Filename là file mẫu (.dwt) hoặc rỗng.
        /// </summary>
        public static string? Problem(string? fileName, bool fileExists, bool fileReadOnly)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return "Bản vẽ chưa có file (Drawing mới chưa lưu lần nào) — lưu tay lần đầu bằng SAVEAS để chọn tên.";
            }

            var extension = Path.GetExtension(fileName);
            if (!string.Equals(extension, ".dwg", StringComparison.OrdinalIgnoreCase))
            {
                return $"Bản vẽ đang gắn với \"{fileName}\" (không phải .dwg — thường là file mẫu của Drawing mới) — lưu tay bằng SAVEAS.";
            }

            if (!fileExists)
            {
                return $"Không thấy file \"{fileName}\" trên đĩa (đã bị xoá/đổi tên?) — lưu tay bằng SAVEAS để tránh ghi nhầm chỗ.";
            }

            if (fileReadOnly)
            {
                return $"File \"{fileName}\" đang ở chế độ chỉ đọc — bỏ thuộc tính read-only hoặc lưu tay sang tên khác.";
            }

            return null;
        }
    }
}
