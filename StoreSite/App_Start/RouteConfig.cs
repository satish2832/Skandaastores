using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace StoreSite
{
    public class RouteConfig
    {
        public static void RegisterRoutes(RouteCollection routes)
        {
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");
            routes.MapRoute(
               name: "Contact",
               url: "Contact",
               defaults: new { controller = "Home", action = "Contact", id = UrlParameter.Optional }
           );
            routes.MapRoute(
              name: "ProductDetails",
              url: "productdetails/{code}/{description}",
              defaults:
              new
              {
                  controller = "Home",
                  action = "ProductDetails",
                  code = UrlParameter.Optional,
                  description = UrlParameter.Optional                 
              }
             );
            routes.MapRoute(
             name: "Admin",
             url: "Admin",
             defaults:
             new
             {
                 controller = "Home",
                 action = "Admin"
             }
            );
            routes.MapRoute(
                name: "Default",
                url: "{controller}/{action}/{id}",
                defaults: new { controller = "Home", action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}
