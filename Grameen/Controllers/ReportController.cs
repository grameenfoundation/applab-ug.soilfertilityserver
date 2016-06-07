using System.Linq;
using System.Web.Mvc;
using Logic;

namespace Grameen.Controllers
{
    public class ReportController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Title = "Report";

            return View();
        }

        public ActionResult ActivityReport()
        {
            ViewBag.Title = "Activity Report";

            return View(OptimizerReport.ActivityReport().OrderByDescending(a => a.dateTime));
        }

        public ActionResult ErrorReport()
        {
            ViewBag.Title = "Error Report";

            return View(OptimizerReport.ErrorReport().OrderByDescending(a => a.dateTime));
        }
    }
}