using Microsoft.Extensions.DependencyInjection;
using WebApiExcelExport.Core.Helpers.ExcelExportServices;

namespace WebApiExcelExport.Core;

public static class CoreServiceRegistrations
{
    public static IServiceCollection AddCoreServices(this IServiceCollection services)
    {
        services.AddScoped<IExcelExportService, ExcelExportService>();
        
        return services;
    }
}