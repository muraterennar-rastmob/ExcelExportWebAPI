using ClosedXML.Excel;

namespace WebApiExcelExport.Core.Helpers.ExcelExportServices;

public class ExcelExportService : IExcelExportService
{
    public MemoryStream ExportToExcel<T>(IEnumerable<T> data, string sheetName)
    {
        using (var workBook = new XLWorkbook())
        {
            var worksheet = workBook.Worksheets.Add(sheetName);

            var properties = typeof(T).GetProperties();

            // Header oluşturma
            for (int i = 0; i < properties.Length; i++)
            {
                worksheet.Cell(1, i + 1).Value = properties[i].Name;
                worksheet.Cell(1, i + 1).Style.Font.Bold = true; // Header'ları kalın yap
            }

            // Data ekleme
            int rowIndex = 2;
            foreach (var item in data)
            {
                for (int colIndex = 0; colIndex < properties.Length; colIndex++)
                {
                    var propType = properties[colIndex].PropertyType;
                    var value = properties[colIndex].GetValue(item);

                    if (value != null)
                    {
                        if (propType == typeof(int) || propType == typeof(long))
                        {
                            worksheet.Cell(rowIndex, colIndex + 1).Value = Convert.ToInt64(value);
                        }
                        else if (propType == typeof(double) || propType == typeof(decimal) || propType == typeof(float))
                        {
                            worksheet.Cell(rowIndex, colIndex + 1).Value = Convert.ToDouble(value);
                        }
                        else if (propType == typeof(DateTime) || propType == typeof(DateTime?))
                        {
                            if (value != null && value is DateTime dateValue && dateValue != DateTime.MinValue)
                            {
                                worksheet.Cell(rowIndex, colIndex + 1).Value = dateValue;
                                worksheet.Cell(rowIndex, colIndex + 1).Style.DateFormat.Format = "yyyy-MM-dd";
                            }
                            else
                            {
                                // Eğer değer null veya DateTime.MinValue ise hücre boş bırakılır
                                worksheet.Cell(rowIndex, colIndex + 1).Value = string.Empty;
                            }
                        }
                        else if (propType == typeof(bool))
                        {
                            worksheet.Cell(rowIndex, colIndex + 1).Value = Convert.ToBoolean(value) ? "Yes" : "No";
                        }
                        else
                        {
                            worksheet.Cell(rowIndex, colIndex + 1).Value = value.ToString();
                        }
                    }
                    else
                    {
                        worksheet.Cell(rowIndex, colIndex + 1).Value = string.Empty; // Null değerler için boş bırak
                    }
                }

                rowIndex++;
            }

            // Otomatik sütun genişliği ayarla
            worksheet.Columns().AdjustToContents();

            // Çıktıyı bir MemoryStream'e yaz
            var stream = new MemoryStream();
            workBook.SaveAs(stream);
            stream.Position = 0; // Stream'in başlangıç noktasına dön

            return stream;
        }
    }
}