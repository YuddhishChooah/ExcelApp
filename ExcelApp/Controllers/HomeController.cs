using BusinessLayer;
using ExcelApp.Models;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;

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
                DataTable mergedData = _excelService.MergeExcel(mainFile, secondFile, model.PrimaryKey1, model.PrimaryKey2, model.GetSelectedColumnsList());

                model.MergedData = mergedData;

                HttpContext.Session.SetString("MergedData", JsonConvert.SerializeObject(mergedData));

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
            var model = new ExcelDataModel();
            return View("ExcelMerge", model);
        }


        public static List<Dictionary<string, object>> DataTableToListOfDictionaries(DataTable dataTable)
        {
            List<Dictionary<string, object>> result = new List<Dictionary<string, object>>();

            foreach (DataRow row in dataTable.Rows)
            {
                Dictionary<string, object> rowDict = new Dictionary<string, object>();

                foreach (DataColumn column in dataTable.Columns)
                {
                    rowDict[column.ColumnName] = row[column];
                }

                result.Add(rowDict);
            }

            return result;
        }

        [HttpGet]
        public IActionResult DownloadMergedFile()
        {
            var mergedDataJson = HttpContext.Session.GetString("MergedData");
            if (!string.IsNullOrEmpty(mergedDataJson))
            {
                var mergedData = JsonConvert.DeserializeObject<DataTable>(mergedDataJson);

                // Create a new Excel file with the merged data
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var excelPackage = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Merged Data");

                    // Load the DataTable into the worksheet
                    worksheet.Cells["A1"].LoadFromDataTable(mergedData, true);

                    // Set the content-type and the file name
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.Headers.Add("content-disposition", $"attachment; filename=MergedData.xlsx");

                    // Write the Excel file to the response stream
                    using (var memoryStream = new MemoryStream())
                    {
                        excelPackage.SaveAs(memoryStream);
                        memoryStream.WriteTo(Response.Body);
                    }
                }
            }
            else
            {
                return NotFound("Merged data not found. Please perform the merge operation first.");
            }

            return new EmptyResult();
        }

        [HttpGet]
        [AllowAnonymous]
        public IActionResult Login()
        {
            return View();
        }

        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Login(LoginViewModel model)
        {
            if (ModelState.IsValid)
            {
                if (ValidateUser(model.Username, model.Password))
                {
                    var claims = new List<Claim>
                    {
                        new Claim(ClaimTypes.Name, model.Username)
                    };

                    var claimsIdentity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);

                    var authProperties = new AuthenticationProperties();

                    await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(claimsIdentity), authProperties);

                    return RedirectToAction("Index");
                }
                else
                {
                    ModelState.AddModelError("", "Invalid username or password.");
                }
            }

            return View(model);
        }

        private bool ValidateUser(string username, string password)
        {
            bool isValid = false;
            string connectionString = "server=localhost;user=root;database=excelappdb;port=3306;";
            string query = "SELECT * FROM excelappdb.users WHERE username=@username AND password=@password";

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                using (MySqlCommand cmd = new MySqlCommand(query, connection))
                {
                    cmd.Parameters.AddWithValue("@username", username);
                    cmd.Parameters.AddWithValue("@password", password);

                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            isValid = true;
                        }
                    }
                }
            }

            return isValid;
        }

    }
}
