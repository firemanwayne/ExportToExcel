using Microsoft.Extensions.DependencyInjection;

namespace Simple.ExportToExcel.UnitTests;

[TestClass]
public class ServiceRegistryExtensionsTests
{
    IServiceCollection _services = null!;

    [TestInitialize]
    public void Setup()
    {
        _services = new ServiceCollection();
        _services.AddExportToExcel();
    }

    // --- IExcelDownloadService registration ---

    [TestMethod]
    public void AddExportToExcel_RegistersIExcelDownloadService()
    {
        var descriptor = _services.FirstOrDefault(d => d.ServiceType == typeof(IExcelDownloadService));

        Assert.IsNotNull(descriptor);
    }

    [TestMethod]
    public void AddExportToExcel_IExcelDownloadService_IsScoped()
    {
        var descriptor = _services.First(d => d.ServiceType == typeof(IExcelDownloadService));

        Assert.AreEqual(ServiceLifetime.Scoped, descriptor.Lifetime);
    }

    [TestMethod]
    public void AddExportToExcel_IExcelDownloadService_ImplementationIsExcelDownloadService()
    {
        var descriptor = _services.First(d => d.ServiceType == typeof(IExcelDownloadService));

        Assert.AreEqual(typeof(ExcelDownloadService), descriptor.ImplementationType);
    }

    // --- IExportToExcel<> registration ---

    [TestMethod]
    public void AddExportToExcel_RegistersIExportToExcel_OpenGeneric()
    {
        var descriptor = _services.FirstOrDefault(d => d.ServiceType == typeof(IExportToExcel<>));

        Assert.IsNotNull(descriptor);
    }

    [TestMethod]
    public void AddExportToExcel_IExportToExcel_IsSingleton()
    {
        var descriptor = _services.First(d => d.ServiceType == typeof(IExportToExcel<>));

        Assert.AreEqual(ServiceLifetime.Singleton, descriptor.Lifetime);
    }

    [TestMethod]
    public void AddExportToExcel_IExportToExcel_ImplementationIsExcelService()
    {
        var descriptor = _services.First(d => d.ServiceType == typeof(IExportToExcel<>));

        Assert.AreEqual(typeof(ExcelService<>), descriptor.ImplementationType);
    }

    // --- Resolution ---

    [TestMethod]
    public void AddExportToExcel_CanResolveIExportToExcel_ForConcreteType()
    {
        var provider = _services.BuildServiceProvider();

        var service = provider.GetService<IExportToExcel<DisplayModel>>();

        Assert.IsNotNull(service);
    }

    [TestMethod]
    public void AddExportToExcel_IExportToExcel_ReturnsSameInstanceAsSingleton()
    {
        var provider = _services.BuildServiceProvider();

        var first  = provider.GetService<IExportToExcel<DisplayModel>>();
        var second = provider.GetService<IExportToExcel<DisplayModel>>();

        Assert.AreSame(first, second);
    }

    [TestMethod]
    public void AddExportToExcel_RegistersExactlyTwoDescriptors()
    {
        Assert.AreEqual(2, _services.Count);
    }
}
