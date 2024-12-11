using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using WebApiExcelExport.Core.Helpers.ExcelExportServices;
using WebApiExcelExport.Domain.Entities;

namespace WebApiExcelExport.API.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ExcelExportController : ControllerBase
{
    private readonly IExcelExportService _excelExportService;

    public ExcelExportController(IExcelExportService excelExportService)
    {
        _excelExportService = excelExportService;
    }


    [HttpGet("v1/export")]
    public IActionResult VersionOneExportExcel()
    {
        var data = new List<dynamic>
        {
            new { Id = 1, Name = "Ali", Age = 30 },
            new { Id = 2, Name = "Ayşe", Age = 25 },
            new { Id = 3, Name = "Fatma", Age = 28 },
        };

        using (var workBook = new XLWorkbook())
        {
            var worksheet = workBook.Worksheets.Add("Sheet1");
            worksheet.Cell(1, 1).Value = "Id";
            worksheet.Cell(1, 2).Value = "Name";
            worksheet.Cell(1, 3).Value = "Age";

            // Data ekleniyor
            for (int i = 0; i < data.Count; i++)
            {
                worksheet.Cell(i + 2, 1).Value = data[i].Id;
                worksheet.Cell(i + 2, 2).Value = data[i].Name;
            }

            // Stream oluşturuluyor ve döndürülüyor
            var stream = new MemoryStream();
            workBook.SaveAs(stream);
            stream.Position = 0; // Okuma pozisyonunu sıfırla

            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            var fileName = "ExportedData.xlsx";
            return File(stream, contentType, fileName);
        }
    }

    [HttpGet("v2/export")]
    public IActionResult VersionTwoExportExcel()
    {
        var data = new List<User>
        {
            new User { Id = 1, Name = "Ali", Age = 30, CreatedDate = DateTime.Now.AddDays(-5) , UpdatedDate = DateTime.Now.AddDays(-3) , DeletedDate = DateTime.Now , IsDeleted = true },
            new User { Id = 2, Name = "Ayşe", Age = 25, CreatedDate = DateTime.Now.AddDays(-2) , UpdatedDate = DateTime.Now.AddDays(-3) , IsDeleted = false },
            new User { Id = 3, Name = "Fatma", Age = 28, CreatedDate = DateTime.Now.AddDays(-3)},
        };

        var stream = _excelExportService.ExportToExcel(data, "Sheet1");

        var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        var fileName = "ExportedData.xlsx";
        return File(stream, contentType, fileName);
    }
}