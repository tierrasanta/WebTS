using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http.Filters;
using System.Web.Mvc;
using System.Web.Mvc.Filters;
using System.Web.Routing;

namespace WebTS2.App_Start
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, Inherited = true, AllowMultiple = true)]
    public class FilterAuth : System.Web.Mvc.ActionFilterAttribute
    {
        public void OnAuthentication(AuthenticationContext filterContext)
        {
        }

        public void OnAuthenticationChallenge(AuthenticationChallengeContext filterContext)
        {
        }

        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            var path = filterContext.HttpContext.Request.CurrentExecutionFilePath;
            HttpSessionStateBase session = filterContext.HttpContext.Session;

            String ControllerName = filterContext.ActionDescriptor.ControllerDescriptor.ControllerName;
            String ActionName = filterContext.ActionDescriptor.ActionName;
            String Method = filterContext.HttpContext.Request.HttpMethod;

            String login = "/Home/Login";
            if (!path.Equals(login) || !path.Contains(login))
            {
                if (HttpContext.Current.Session["Usuario"] == null)
                {
                    filterContext.Result = new RedirectToRouteResult(
                        new RouteValueDictionary{
                            { "area", "" },
                            { "controller", "Home" },
                            { "action", "Login" }
                        });
                    filterContext.Result.ExecuteResult(filterContext.Controller.ControllerContext);
                }
            }
        }
    }
}