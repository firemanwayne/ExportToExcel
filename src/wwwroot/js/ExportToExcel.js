
/**
 * 
 * @param {string} FileName
 * @param {string} FileContent
 * @param {string} ContentType
 * @param {any} FileAbsoluteUri
 */
function ExportHelper(FileName, FileContent, ContentType, FileAbsoluteUri) {

    function GenerateFileFromBytes(FileName, FileContent, ContentType) {
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
    }

    function GenerateFileFromUri(FileName, FileAbsoluteUri) {
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
    }

    if (FileAbsoluteUri === null) {
        GenerateFileFromBytes(FileName, FileContent, ContentType);
    }

    if (FileAbsoluteUri !== null) {
        GenerateFileFromUri(FileName, FileAbsoluteUri);
    }   
}

export function ExportFile(FileName, FileContent, ContentType) {
    ExportHelper(FileName, FileContent, ContentType, null);
}

export function ExportFileToUri(FileName, FileAbsoluteUri) {
    ExportHelper(FileName, null, null, FileAbsoluteUri);
}