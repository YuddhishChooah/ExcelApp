using Microsoft.AspNetCore.Mvc;

namespace ExcelApp.Controllers
{
    public class ExcelController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
