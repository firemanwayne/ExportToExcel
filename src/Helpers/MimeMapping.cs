using System;
using System.Collections.Generic;
using System.Linq;

namespace Simple.ExportToExcel;

internal class MimeMapping
{
    private MimeMapping() { }

    public static IDictionary<string, Extension> Extensions = new Dictionary<string, Extension>
    {
        {".acx", new Extension(".acx", "application/internet-property-stream")},
        {".ai", new Extension(".ai", "application/postscript")},
        {".aif", new Extension(".aif", "audio/x-aiff")},
        {".aiff", new Extension(".aiff", "audio/aiff")},
        {".axs", new Extension(".axs", "application/olescript")},
        {".aifc", new Extension(".aifc", "audio/aiff")},
        {".asr", new Extension(".asr", "video/x-ms-asf")},
        {".avi", new Extension(".avi", "video/x-msvideo")},
        {".asf", new Extension(".asf", "video/x-ms-asf")},
        {".au", new Extension(".au", "audio/basic")},
        {".application", new Extension(".application", "application/x-ms-application")},
        {".bin", new Extension(".bin", "application/octet-stream")},
        {".bas", new Extension(".bas", "text/plain")},
        {".bcpio", new Extension(".bcpio", "application/x-bcpio")},
        {".bmp", new Extension(".bmp", "image/bmp")},
        {".cdf", new Extension(".cdf", "application/x-cdf")},
        {".cat", new Extension(".cat", "application/vndms-pkiseccat")},
        {".crt", new Extension(".crt", "application/x-x509-ca-cert")},
        {".c", new Extension(".c", "text/plain")},
        {".css", new Extension(".css", "text/css")},
        {".cer", new Extension(".cer", "application/x-x509-ca-cert")},
        {".crl", new Extension(".crl", "application/pkix-crl")},
        {".cmx", new Extension(".cmx", "image/x-cmx")},
        {".csh", new Extension(".csh", "application/x-csh")},
        {".cod", new Extension(".cod", "image/cis-cod")},
        {".cpio", new Extension(".cpio", "application/x-cpio")},
        {".clp", new Extension(".clp", "application/x-msclip")},
        {".crd", new Extension(".crd", "application/x-mscardfile")},
        {".deploy", new Extension(".deploy", "application/octet-stream")},
        {".dll", new Extension(".dll", "application/x-msdownload")},
        {".dot", new Extension(".dot", "application/msword")},
        {".doc", new Extension(".doc", "application/msword")},
        {".dvi", new Extension(".dvi", "application/x-dvi")},
        {".dir", new Extension(".dir", "application/x-director")},
        {".dxr", new Extension(".dxr", "application/x-director")},
        {".der", new Extension(".der", "application/x-x509-ca-cert")},
        {".dib", new Extension(".dib", "image/bmp")},
        {".dcr", new Extension(".dcr", "application/x-director")},
        {".disco", new Extension(".disco", "text/xml")},
        {".exe", new Extension(".exe", "application/octet-stream")},
        {".etx", new Extension(".etx", "text/x-setext")},
        {".evy", new Extension(".evy", "application/envoy")},
        {".eml", new Extension(".eml", "message/rfc822")},
        {".eps", new Extension(".eps", "application/postscript")},
        {".flr", new Extension(".flr", "x-world/x-vrml")},
        {".fif", new Extension(".fif", "application/fractals")},
        {".gtar", new Extension(".gtar", "application/x-gtar")},
        {".gif", new Extension(".gif", "image/gif")},
        {".gz", new Extension(".gz", "application/x-gzip")},
        {".hta", new Extension(".hta", "application/hta")},
        {".htc", new Extension(".htc", "text/x-component")},
        {".htt", new Extension(".htt", "text/webviewhtml")},
        {".h", new Extension(".h", "text/plain")},
        {".hdf", new Extension(".hdf", "application/x-hdf")},
        {".hlp", new Extension(".hlp", "application/winhlp")},
        {".html", new Extension(".html", "text/html")},
        {".htm", new Extension(".htm", "text/html")},
        {".hqx", new Extension(".hqx", "application/mac-binhex40")},
        {".isp", new Extension(".isp", "application/x-internet-signup")},
        {".iii", new Extension(".iii", "application/x-iphone")},
        {".ief", new Extension(".ief", "image/ief")},
        {".ivf", new Extension(".ivf", "video/x-ivf")},
        {".ins", new Extension(".ins", "application/x-internet-signup")},
        {".ico", new Extension(".ico", "image/x-icon")},
        {".jpg", new Extension(".jpg", "image/jpeg")},
        {".jfif", new Extension(".jfif", "image/pjpeg")},
        {".jpe", new Extension(".jpe", "image/jpeg")},
        {".jpeg", new Extension(".jpeg", "image/jpeg")},
        {".js", new Extension(".js", "application/x-javascript")},
        {".lsx", new Extension(".lsx", "video/x-la-asf")},
        {".latex", new Extension(".latex", "application/x-latex")},
        {".lsf", new Extension(".lsf", "video/x-la-asf")},
        {".manifest", new Extension(".manifest", "application/x-ms-manifest")},
        {".mhtml", new Extension(".mhtml", "message/rfc822")},
        {".mny", new Extension(".mny", "application/x-msmoney")},
        {".mht", new Extension(".mht", "message/rfc822")},
        {".mid", new Extension(".mid", "audio/mid")},
        {".mpv2", new Extension(".mpv2", "video/mpeg")},
        {".man", new Extension(".man", "application/x-troff-man")},
        {".mvb", new Extension(".mvb", "application/x-msmediaview")},
        {".mpeg", new Extension(".mpeg", "video/mpeg")},
        {".m3u", new Extension(".m3u", "audio/x-mpegurl")},
        {".mdb", new Extension(".mdb", "application/x-msaccess")},
        {".mpp", new Extension(".mpp", "application/vnd.ms-project")},
        {".m1v", new Extension(".m1v", "video/mpeg")},
        {".mpa", new Extension(".mpa", "video/mpeg")},
        {".me", new Extension(".me", "application/x-troff-me")},
        {".m13", new Extension(".m13", "application/x-msmediaview")},
        {".movie", new Extension(".movie", "video/x-sgi-movie")},
        {".m14", new Extension(".m14", "application/x-msmediaview")},
        {".mpe", new Extension(".mpe", "video/mpeg")},
        {".mp2", new Extension(".mp2", "video/mpeg")},
        {".mov", new Extension(".mov", "video/quicktime")},
        {".mp3", new Extension(".mp3", "audio/mpeg")},
        {".mpg", new Extension(".mpg", "video/mpeg")},
        {".ms", new Extension(".ms", "application/x-troff-ms")},
        {".nc", new Extension(".nc", "application/x-netcdf")},
        {".nws", new Extension(".nws", "message/rfc822")},
        {".oda", new Extension(".oda", "application/oda")},
        {".ods", new Extension(".ods", "application/oleobject")},
        {".pmc", new Extension(".pmc", "application/x-perfmon")},
        {".p7r", new Extension(".p7r", "application/x-pkcs7-certreqresp")},
        {".p7b", new Extension(".p7b", "application/x-pkcs7-certificates")},
        {".p7s", new Extension(".p7s", "application/pkcs7-signature")},
        {".pmw", new Extension(".pmw", "application/x-perfmon")},
        {".ps", new Extension(".ps", "application/postscript")},
        {".p7c", new Extension(".p7c", "application/pkcs7-mime")},
        {".pbm", new Extension(".pbm", "image/x-portable-bitmap")},
        {".ppm", new Extension(".ppm", "image/x-portable-pixmap")},
        {".pub", new Extension(".pub", "application/x-mspublisher")},
        {".pnm", new Extension(".pnm", "image/x-portable-anymap")},
        {".pml", new Extension(".pml", "application/x-perfmon")},
        {".p10", new Extension(".p10", "application/pkcs10")},
        {".pfx", new Extension(".pfx", "application/x-pkcs12")},
        {".p12", new Extension(".p12", "application/x-pkcs12")},
        {".pdf", new Extension(".pdf", "application/pdf")},
        {".pps", new Extension(".pps", "application/vnd.ms-powerpoint")},
        {".p7m", new Extension(".p7m", "application/pkcs7-mime")},
        {".pko", new Extension(".pko", "application/vndms-pkipko")},
        {".ppt", new Extension(".ppt", "application/vnd.ms-powerpoint")},
        {".pmr", new Extension(".pmr", "application/x-perfmon")},
        {".pma", new Extension(".pma", "application/x-perfmon")},
        {".pot", new Extension(".pot", "application/vnd.ms-powerpoint")},
        {".prf", new Extension(".prf", "application/pics-rules")},
        {".pgm", new Extension(".pgm", "image/x-portable-graymap")},
        {".qt", new Extension(".qt", "video/quicktime")},
        {".ra", new Extension(".ra", "audio/x-pn-realaudio")},
        {".rgb", new Extension(".rgb", "image/x-rgb")},
        {".ram", new Extension(".ram", "audio/x-pn-realaudio")},
        {".rmi", new Extension(".rmi", "audio/mid")},
        {".ras", new Extension(".ras", "image/x-cmu-raster")},
        {".roff", new Extension(".roff", "application/x-troff")},
        {".rtf", new Extension(".rtf", "application/rtf")},
        {".rtx", new Extension(".rtx", "text/richtext")},
        {".sv4crc", new Extension(".sv4crc", "application/x-sv4crc")},
        {".spc", new Extension(".spc", "application/x-pkcs7-certificates")},
        {".setreg", new Extension(".setreg", "application/set-registration-initiation")},
        {".snd", new Extension(".snd", "audio/basic")},
        {".stl", new Extension(".stl", "application/vndms-pkistl")},
        {".setpay", new Extension(".setpay", "application/set-payment-initiation")},
        {".stm", new Extension(".stm", "text/html")},
        {".shar", new Extension(".shar", "application/x-shar")},
        {".sh", new Extension(".sh", "application/x-sh")},
        {".sit", new Extension(".sit", "application/x-stuffit")},
        {".spl", new Extension(".spl", "application/futuresplash")},
        {".sct", new Extension(".sct", "text/scriptlet")},
        {".scd", new Extension(".scd", "application/x-msschedule")},
        {".sst", new Extension(".sst", "application/vndms-pkicertstore")},
        {".src", new Extension(".src", "application/x-wais-source")},
        {".sv4cpio", new Extension(".sv4cpio", "application/x-sv4cpio")},
        {".tex", new Extension(".tex", "application/x-tex")},
        {".tgz", new Extension(".tgz", "application/x-compressed")},
        {".t", new Extension(".t", "application/x-troff")},
        {".tar", new Extension(".tar", "application/x-tar")},
        {".tr", new Extension(".tr", "application/x-troff")},
        {".tif", new Extension(".tif", "image/tiff")},
        {".txt", new Extension(".txt", "text/plain")},
        {".texinfo", new Extension(".texinfo", "application/x-texinfo")},
        {".trm", new Extension(".trm", "application/x-msterminal")},
        {".tiff", new Extension(".tiff", "image/tiff")},
        {".tcl", new Extension(".tcl", "application/x-tcl")},
        {".texi", new Extension(".texi", "application/x-texinfo")},
        {".tsv", new Extension(".tsv", "text/tab-separated-values")},
        {".ustar", new Extension(".ustar", "application/x-ustar")},
        {".uls", new Extension(".uls", "text/iuls")},
        {".vcf", new Extension(".vcf", "text/x-vcard")},
        {".wps", new Extension(".wps", "application/vnd.ms-works")},
        {".wav", new Extension(".wav", "audio/wav")},
        {".wrz", new Extension(".wrz", "x-world/x-vrml")},
        {".wri", new Extension(".wri", "application/x-mswrite")},
        {".wks", new Extension(".wks", "application/vnd.ms-works")},
        {".wmf", new Extension(".wmf", "application/x-msmetafile")},
        {".wcm", new Extension(".wcm", "application/vnd.ms-works")},
        {".wrl", new Extension(".wrl", "x-world/x-vrml")},
        {".wdb", new Extension(".wdb", "application/vnd.ms-works")},
        {".wsdl", new Extension(".wsdl", "text/xml")},
        {".xml", new Extension(".xml", "text/xml")},
        {".xlm", new Extension(".xlm", "application/vnd.ms-excel")},
        {".xaf", new Extension(".xaf", "x-world/x-vrml")},
        {".xla", new Extension(".xla", "application/vnd.ms-excel")},
        {".xls", new Extension(".xls", "application/vnd.ms-excel")},
        {".xof", new Extension(".xof", "x-world/x-vrml")},
        {".xlt", new Extension(".xlt", "application/vnd.ms-excel")},
        {".xlc", new Extension(".xlc", "application/vnd.ms-excel")},
        {".xsl", new Extension(".xsl", "text/xml")},
        {".xbm", new Extension(".xbm", "image/x-xbitmap")},
        {".xlw", new Extension(".xlw", "application/vnd.ms-excel")},
        {".xpm", new Extension(".xpm", "image/x-xpixmap")},
        {".xwd", new Extension(".xwd", "image/x-xwindowdump")},
        {".xsd", new Extension(".xsd", "text/xml")},
        {".z", new Extension(".z", "application/x-compress")},
        {".zip", new Extension(".zip", "application/x-zip-compressed")},
        {".*", new Extension(".*", "application/octet-stream")}
    };

    public static IDictionary<string, string> GetAllMimeTypes()
    {
        Dictionary<string, string> OrderedDictionary = new();
        foreach (KeyValuePair<string, Extension> item in Extensions.OrderBy(m => m.Value))
        {
            Extension extension = item.Value;
            OrderedDictionary.TryAdd(item.Key, $"Type: {GetMediaType(extension.Value)} Ext: {item.Key}");
        }
        return OrderedDictionary;
    }

    public static string GetMimeMappingByExtension(string fileExtension)
    {
        if (Extensions.ContainsKey(fileExtension))
        {
            return Extensions[fileExtension].Value;
        }
        else
        {
            return Extensions[".*"].Value;
        }
    }

    public static string GetMimeFromFileName(string FileName)
    {
        if (!IsExtensionMissing(FileName))
        {
            string[] ValueArray = FileName.Split('.');
            if (ValueArray.Length > 1)
            {
                if (Extensions.ContainsKey($".{ValueArray[1]}"))
                {
                    return Extensions[$".{ValueArray[1]}"].Value;
                }
                else
                {
                    return Extensions[".*"].Value;
                }
            }

            return "";
        }
        else
        {
            throw new FileNameMissingExtensionException();
        }
    }

    public static bool IsExtensionMissing(string FileName)
    {
        string[] ValueArray = FileName.Split('.');
        if (ValueArray.Length > 1)
        {
            return false;
        }
        else
        {
            return true;
        }
    }

    public static List<string> GetMimeMappingsByType(string MimeTypeName)
    {
        List<string> Mimes = new();
        foreach (KeyValuePair<string, Extension> item in Extensions)
        {
            if (Extensions.Values.Select(m => m.Value).Contains(MimeTypeName))
            {
                Mimes.Add(item.Key);
            }
        }

        return Mimes;
    }

    static string GetMediaType(string MediaTypeName)
    {
        string[] MediaTypeArray = MediaTypeName.Split('/');
        return MediaTypeArray[0];
    }

    static string GetMediaName(string MediaTypeName)
    {
        return MediaTypeName.Split('/').Length > 0 ?
            MediaTypeName.Split('/')[1] : "Unknown";
    }
}

public struct Extension
{
    public string Key { get; set; }

    public string Value { get; set; }

    public Extension(string Key, string Value)
    {
        this.Key = Key;
        this.Value = Value;
    }
}