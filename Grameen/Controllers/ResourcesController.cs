using System.IO;
using System.Net.Mime;
using System.Web;
using System.Web.Mvc;

namespace Grameen.Controllers
{
    public class ResourcesController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Title = "Resources";

            return View();
        }

        public ActionResult apkLink()
        {
            ViewBag.Title = "Resources";

            var apkFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\fot.apk");

            // Download the file

            var filename = apkFile.Name;
            var filepath = apkFile.FullName;
            var filedata = System.IO.File.ReadAllBytes(filepath);
            var contentType = MimeMapping.GetMimeMapping(filepath);

            var cd = new ContentDisposition
            {
                FileName = filename,
                Inline = true
            };

            Response.AppendHeader("Content-Disposition", cd.ToString());

            return File(filedata, contentType);
        }

        public ActionResult userManualLink()
        {
            ViewBag.Title = "User Manual";

            var apkFile = new FileInfo(@"C:\inetpub\wwwroot\Temp\User_manual.docx");

            // Download the file

            var filename = apkFile.Name;
            var filepath = apkFile.FullName;
            var filedata = System.IO.File.ReadAllBytes(filepath);
            var contentType = MimeMapping.GetMimeMapping(filepath);

            var cd = new ContentDisposition
            {
                FileName = filename,
                Inline = true
            };

            Response.AppendHeader("Content-Disposition", cd.ToString());

            return File(filedata, contentType);
        }
    }
}