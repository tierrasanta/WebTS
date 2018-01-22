using System.Web;
using System.Web.Mvc;
using WebTS2.App_Start;

namespace WebTS2
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
            filters.Add(new FilterAuth());
        }
    }
}
