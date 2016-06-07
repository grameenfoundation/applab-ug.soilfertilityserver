using System.Web.Http;

namespace Grameen
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Web API configuration and services

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute("DefaultApi", "grameen/{controller}/{id}", new {id = RouteParameter.Optional}
                );
        }
    }
}