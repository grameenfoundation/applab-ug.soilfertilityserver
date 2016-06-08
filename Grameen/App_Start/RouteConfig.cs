using System.Web.Mvc;
using System.Web.Routing;

namespace Grameen
{
    public class RouteConfig
    {
        public static void RegisterRoutes(RouteCollection routes)
        {
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");

            routes.MapRoute("Default", "{controller}/{action}/{id}",
                new {controller = "Resources", action = "Index", id = UrlParameter.Optional}
                );
        }
    }
}