using Microsoft.JSInterop;

namespace Simple.ExportToExcel.UnitTests;

// ---------------------------------------------------------------------------
// Test doubles
// ---------------------------------------------------------------------------

file sealed class FakeJSObjectReference : IJSObjectReference
{
    public List<(string Identifier, object?[]? Args)> Calls { get; } = [];
    public bool DisposeCalled { get; private set; }

    public ValueTask<TValue> InvokeAsync<TValue>(string identifier, object?[]? args)
    {
        Calls.Add((identifier, args));
        return ValueTask.FromResult(default(TValue)!);
    }

    public ValueTask<TValue> InvokeAsync<TValue>(string identifier, CancellationToken cancellationToken, object?[]? args)
    {
        Calls.Add((identifier, args));
        return ValueTask.FromResult(default(TValue)!);
    }

    public ValueTask DisposeAsync()
    {
        DisposeCalled = true;
        return ValueTask.CompletedTask;
    }
}

file sealed class FakeJSRuntime : IJSRuntime
{
    public FakeJSObjectReference Module { get; } = new();

    public ValueTask<TValue> InvokeAsync<TValue>(string identifier, object?[]? args)
        => ValueTask.FromResult((TValue)(object)Module);

    public ValueTask<TValue> InvokeAsync<TValue>(string identifier, CancellationToken cancellationToken, object?[]? args)
        => ValueTask.FromResult((TValue)(object)Module);
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

[TestClass]
public class ExcelDownloadServiceTests
{
    [TestMethod]
    public async Task ExportFile_WithLocalResponse_InvokesLocalExportFunction()
    {
        var js = new FakeJSRuntime();
        var service = new ExcelDownloadService(js);
        var response = new UploadFileLocalResponse
        {
            FileName = "report.xlsx",
            ContentType = ExcelConstants.Excel,
            FileContent = "base64data",
        };

        await service.ExportFile(response);

        Assert.AreEqual(1, js.Module.Calls.Count);
        Assert.AreEqual("ExportFile", js.Module.Calls[0].Identifier);
    }

    [TestMethod]
    public async Task ExportFile_WithLocalResponse_PassesCorrectArguments()
    {
        var js = new FakeJSRuntime();
        var service = new ExcelDownloadService(js);
        var response = new UploadFileLocalResponse
        {
            FileName = "report.xlsx",
            ContentType = ExcelConstants.Excel,
            FileContent = "abc123",
        };

        await service.ExportFile(response);

        var args = js.Module.Calls[0].Args!;
        Assert.AreEqual("report.xlsx", args[0]);
        Assert.AreEqual("abc123",      args[1]);
        Assert.AreEqual(ExcelConstants.Excel, args[2]);
    }

    [TestMethod]
    public async Task ExportFile_WithUriResponse_InvokesUriExportFunction()
    {
        var js = new FakeJSRuntime();
        var service = new ExcelDownloadService(js);
        var response = new UploadFileResponse
        {
            FileName = "report.xlsx",
            ContentType = ExcelConstants.Excel,
            FileUri = new Uri("https://storage.example.com/report.xlsx"),
        };

        await service.ExportFile(response);

        Assert.AreEqual(1, js.Module.Calls.Count);
        Assert.AreEqual("ExportFileToUri", js.Module.Calls[0].Identifier);
    }

    [TestMethod]
    public async Task ExportFile_WithUriResponse_PassesCorrectArguments()
    {
        var js = new FakeJSRuntime();
        var service = new ExcelDownloadService(js);
        var uri = new Uri("https://storage.example.com/report.xlsx");
        var response = new UploadFileResponse
        {
            FileName = "report.xlsx",
            ContentType = ExcelConstants.Excel,
            FileUri = uri,
        };

        await service.ExportFile(response);

        var args = js.Module.Calls[0].Args!;
        Assert.AreEqual("report.xlsx", args[0]);
        Assert.AreEqual(uri,           args[1]);
    }

    [TestMethod]
    public async Task ExportFile_CalledTwice_InvokesFunctionTwice()
    {
        var js = new FakeJSRuntime();
        var service = new ExcelDownloadService(js);
        var response = new UploadFileLocalResponse
        {
            FileName = "a.xlsx",
            ContentType = ExcelConstants.Excel,
            FileContent = "x",
        };

        await service.ExportFile(response);
        await service.ExportFile(response);

        Assert.AreEqual(2, js.Module.Calls.Count);
    }

    [TestMethod]
    public async Task DisposeAsync_WhenModuleNotYetLoaded_DoesNotDisposeModule()
    {
        var js = new FakeJSRuntime();
        var service = new ExcelDownloadService(js);

        await service.DisposeAsync();

        Assert.IsFalse(js.Module.DisposeCalled);
    }

    [TestMethod]
    public async Task DisposeAsync_AfterModuleLoaded_DisposesModule()
    {
        var js = new FakeJSRuntime();
        var service = new ExcelDownloadService(js);
        // Trigger module load
        await service.ExportFile(new UploadFileLocalResponse
        {
            FileName = "x.xlsx",
            ContentType = ExcelConstants.Excel,
            FileContent = "x",
        });

        await service.DisposeAsync();

        Assert.IsTrue(js.Module.DisposeCalled);
    }
}
