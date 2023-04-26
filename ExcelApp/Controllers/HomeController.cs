using BusinessLayer;
using ExcelApp.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Data;
using System.Diagnostics;

namespace ExcelApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IExcelService _excelService;

        public HomeController(ILogger<HomeController> logger, IExcelService excelService)
        {
            _logger = logger;
            _excelService = excelService;
        }

        [HttpPost]
        public IActionResult ExcelMerge(ExcelDataModel model, IFormFile mainFile, IFormFile secondFile)
        {
            if (mainFile != null && secondFile != null)
            {
                DataTable mergedData = _excelService.MergeExcel(mainFile, secondFile, model.PrimaryKey1, model.PrimaryKey2, model.SelectedColumn);

                model.MergedData = mergedData;
            }
            else
            {
                ModelState.AddModelError("", "Please upload both files.");
            }

            return View(model);
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }
    }
}