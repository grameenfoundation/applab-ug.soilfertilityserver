using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using Grameen.Models;
using Optimize;

namespace Grameen.Controllers
{
    public class ReportController : Controller
    {
        private ApplicationDbContext database = new ApplicationDbContext();

        public ActionResult Index()
        {
            ViewBag.Title = "Report";

            return View();
        }

        public ActionResult ActivityReport()
        {
            ViewBag.Title = "Activity Report";

            var activities = database.Activities.ToList();
                 
            var model = new List<ActivityView>();
            foreach (Activity activity in activities)
            {
                try
                {
                    model.Add(new ActivityView
                    {
                        Date = activity.DateTime,
                        Calculation = new JavaScriptSerializer().Deserialize<Calc>(activity.Calculation.ToString())
                    });
                }
                catch (Exception e)
                {
                    //database.Activities.Remove(activity);
                    //database.SaveChanges();
                   // throw;
                }
            }
            return View(model.OrderByDescending(a => a.Date));
        }

        public ActionResult ErrorReport()
        {
            ViewBag.Title = "Error Report";

            var errors = database.Errors.ToList().Select(a =>
                new ErrorView
                {
                    Date =  a.DateTime ,
                    Calculation = new JavaScriptSerializer().Deserialize<Calc>(a.Calculation.ToString()),
                    error = a.error
                }
                );

            return View(errors.OrderByDescending(a => a.Date));
        }
    }
}