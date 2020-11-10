

class FileSaverHelper {

    constructor(FileName: string, FileContent: string, ContentType: string, FileAbsoluteUri: string) {
        if (FileAbsoluteUri === null) {
            this.GenerateFileFromBytes(FileName, FileContent, ContentType);
        }

        if (FileAbsoluteUri !== null) {
            this.GenerateFileFromUri(FileName, FileAbsoluteUri);
        }
    }

    private GenerateFileFromBytes = (FileName: string, FileContent: string, ContentType: string) => {
        try {
            const link = document.createElement('a');
            link.download = FileName;
            link.href = "data:" + ContentType + ";base64," + encodeURIComponent(FileContent);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

        } catch (e) {
            console.log('Error: ' + e);
        }
    }
    private GenerateFileFromUri = (FileName: string, FileAbsoluteUri: string) => {
        try {
            const link = document.createElement('a');
            link.download = FileName;
            link.href = FileAbsoluteUri;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

        } catch (e) {
            console.log('Error: ' + e);
        }
    }
}

function ExportFile(FileName: string, FileContent: string, ContentType: string): void {
    new FileSaverHelper(FileName, FileContent, ContentType, null);
}
function ExportFileToUri(FileName: string, FileAbsoluteUri: string): void {
    new FileSaverHelper(FileName, null, null, FileAbsoluteUri);
}