using WebApiExcelExport.Domain.Common;

namespace WebApiExcelExport.Domain.Entities;

public class User:BaseEntity
{
    public int Id { get; set; }
    public string Name { get; set; }
    public int Age { get; set; }
}