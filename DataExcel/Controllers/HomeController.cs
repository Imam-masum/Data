using DataExcel.Models;
using EntityFramework.BulkExtensions;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Diagnostics;

namespace DataExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly AppDbContext context;
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger, AppDbContext context)
        {
            _logger = logger;
            this.context = context;
        }

        public IActionResult Index()
        {
            return View();
        }
        public async Task<List<Country>> Import(IFormFile formFile)
        {
            var list=new List<Country>();
            using (var stream=new MemoryStream())
            {
                await formFile.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream)) 
                {
                    ExcelWorksheet excelWorksheet = package.Workbook.Worksheets[0];
                    var rowcount = excelWorksheet.Dimension.Rows;
                    for (int row = 2; row <= rowcount; row++)
                    {
                        list.Add(new Country
                        {
                            Id = excelWorksheet.Cells[row, 1].Value.ToString().Trim(),
                            CountryName = excelWorksheet.Cells[row,2].Value.ToString().Trim(),
                            TwoCode = excelWorksheet.Cells[row,3].Value.ToString().Trim(),
                            TthreeCode=excelWorksheet.Cells.Value.ToString().Trim()
                            
                            
                        });
                       
                    }

                }
                //formFile = new FormFile(list);
                context.Add(stream);
                context.SaveChanges();
            }
            return list;
        }
        //public List<Country> SaveCountry(List<Country> countries)
        //{
        //    context.BulkInsert(countries);
        //    return countries;
        //}
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}