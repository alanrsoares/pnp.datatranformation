using System.Web;
using System.Web.Mvc;

namespace Vtex.Practices.DataTransformation.SampleWebApp
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}