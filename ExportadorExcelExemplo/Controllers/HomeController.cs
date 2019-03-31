using ExportadorExcelExemplo.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace ExportadorExcelExemplo.Controllers
{
    public class HomeController : Controller
    {
        IHostingEnvironment _hostingEnvironment;

        public HomeController(IHostingEnvironment  hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpGet]
        public async Task<IActionResult> Export()
        {

            List<Pessoa> pessoas = new List<Pessoa>()
            {
                new Pessoa (1, "Fulano", "1234443", "232323"),
                new Pessoa (3, "Fulano", "23223213", "5454"),
                new Pessoa (4, "Fulano", "132131231", "5454"),
                new Pessoa (5, "Fulano", "453453", "545454"),
            };

            string sWebRootFolder = _hostingEnvironment.WebRootPath;
            string sFileName = @"demo.xlsx";
            string URL = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, sFileName);
            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
            var memory = new MemoryStream();
            using (var fs = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook;
                workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Folha");

                int rowCount = 1;
                IRow row = excelSheet.CreateRow(0);
                row.CreateCell(0).SetCellValue("Id");
                row.CreateCell(1).SetCellValue("Nome");
                row.CreateCell(2).SetCellValue("CPF");
                row.CreateCell(3).SetCellValue("TELEFONE");


                foreach (var pessoa in pessoas)
                {
                    row = excelSheet.CreateRow(rowCount);
                    row.CreateCell(0).SetCellValue(pessoa.Id);
                    row.CreateCell(1).SetCellValue(pessoa.Nome);
                    row.CreateCell(2).SetCellValue(pessoa.Cpf);
                    row.CreateCell(3).SetCellValue(pessoa.Telefone);

                    rowCount++;
                }

                workbook.Write(fs);
            }
            using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;
            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", sFileName);
        }

    }
}
