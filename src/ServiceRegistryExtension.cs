namespace Microsoft.Extensions.DependencyInjection;

/// <summary>
/// Extension methods for <see cref="IServiceCollection"/> to register ExportToExcel services.
/// </summary>
public static class ServiceBusConfigurationExtension
{
    /// <summary>
    /// Registers the ExportToExcel services required for spreadsheet generation and browser download.
    /// Adds <see cref="Simple.ExportToExcel.IExcelDownloadService"/> as scoped and
    /// <see cref="Simple.ExportToExcel.IExportToExcel{T}"/> as a singleton open-generic.
    /// </summary>
    /// <param name="services">The <see cref="IServiceCollection"/> to add services to.</param>
    public static void AddExportToExcel(this IServiceCollection services)
    {
        services.AddScoped<IExcelDownloadService, ExcelDownloadService>();

        services.AddSingleton(typeof(IExportToExcel<>), typeof(ExcelService<>));
    }
}