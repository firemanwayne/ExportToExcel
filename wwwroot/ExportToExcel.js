var FileSaverHelper = /** @class */ (function () {
    function FileSaverHelper(FileName, FileContent, ContentType, FileAbsoluteUri) {
        this.GenerateFileFromBytes = function (FileName, FileContent, ContentType) {
            try {
                var link = document.createElement('a');
                link.download = FileName;
                link.href = "data:" + ContentType + ";base64," + encodeURIComponent(FileContent);
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            }
            catch (e) {
                console.log('Error: ' + e);
            }
        };
        this.GenerateFileFromUri = function (FileName, FileAbsoluteUri) {
            try {
                var link = document.createElement('a');
                link.download = FileName;
                link.href = FileAbsoluteUri;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            }
            catch (e) {
                console.log('Error: ' + e);
            }
        };
        if (FileAbsoluteUri === null) {
            this.GenerateFileFromBytes(FileName, FileContent, ContentType);
        }
        if (FileAbsoluteUri !== null) {
            this.GenerateFileFromUri(FileName, FileAbsoluteUri);
        }
    }
    return FileSaverHelper;
}());
function ExportFile(FileName, FileContent, ContentType) {
    new FileSaverHelper(FileName, FileContent, ContentType, null);
}
function ExportFileToUri(FileName, FileAbsoluteUri) {
    new FileSaverHelper(FileName, null, null, FileAbsoluteUri);
}
//# sourceMappingURL=ExportToExcel.js.map