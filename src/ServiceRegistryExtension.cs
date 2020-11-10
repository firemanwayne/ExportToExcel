using ExportToExcel.Infrastructure;
using Microsoft.Extensions.DependencyInjection;

namespace ExportToExcel
{
    public static class ServiceBusConfigurationExtension
    {
        /// <summary>
        /// Registers services used by Export To Excel
        /// </summary>
        /// <param name="services"></param>
        /// <returns></returns>
        public static void AddExportToExcel(this IServiceCollection services)
        {
            services.AddScoped<ExcelDownloadService>();

            services.AddSingleton(typeof(IExportToExcel<>), typeof(ExcelService<>));
        }
    }
}