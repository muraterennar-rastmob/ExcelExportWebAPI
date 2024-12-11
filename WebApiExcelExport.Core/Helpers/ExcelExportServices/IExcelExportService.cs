namespace WebApiExcelExport.Core.Helpers.ExcelExportServices;

public interface IExcelExportService
{
    MemoryStream ExportToExcel<T>(IEnumerable<T> data, string sheetName);
}