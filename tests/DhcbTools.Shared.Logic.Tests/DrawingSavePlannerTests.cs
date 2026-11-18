using DhcbTools.Shared.Logic.Cad;
using Xunit;

namespace DhcbTools.Shared.Logic.Tests;

/// <summary>Điều kiện của lệnh DrawingSave: chỉ lưu đè đúng file .dwg đang tồn tại và ghi được.</summary>
public class DrawingSavePlannerTests
{
    [Fact]
    public void FileDwgTonTaiGhiDuoc_LuuDuoc()
    {
        Assert.Null(DrawingSavePlanner.Problem(@"C:\DuAn\Tang ham 2.dwg", fileExists: true, fileReadOnly: false));
        Assert.Null(DrawingSavePlanner.Problem(@"C:\DuAn\HOA.DWG", fileExists: true, fileReadOnly: false));
    }

    [Theory]
    [InlineData(null, true, false, "chưa có file")]
    [InlineData("  ", true, false, "chưa có file")]
    [InlineData(@"C:\Program Files\Autodesk\AutoCAD 2026\Template\acad.dwt", true, false, "không phải .dwg")]
    [InlineData(@"C:\DuAn\ban-ve.dxf", true, false, "không phải .dwg")]
    [InlineData(@"C:\DuAn\da-xoa.dwg", false, false, "Không thấy file")]
    [InlineData(@"C:\DuAn\chi-doc.dwg", true, true, "chỉ đọc")]
    public void TruongHopKhongLuu_BaoLyDo(string? fileName, bool exists, bool readOnly, string expected)
    {
        var problem = DrawingSavePlanner.Problem(fileName, exists, readOnly);
        Assert.NotNull(problem);
        Assert.Contains(expected, problem);
    }
}
