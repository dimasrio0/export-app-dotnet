using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using Rotativa.AspNetCore;

namespace ExportApp.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;

    public HomeController(ILogger<HomeController> logger)
    {
        _logger = logger;
    }

    public IActionResult Index()
    {
        return View();
    }

    public IActionResult ExportPdf()
    {
        var model = new List<string> { "Satu", "Dua", "Tiga" };

        return new ViewAsPdf("PdfView", model)
        {
            FileName = "Contoh-PDF.pdf",
            PageSize = Rotativa.AspNetCore.Options.Size.A4,
        };
    }

    public IActionResult ExportExcel()
    {
        ExcelPackage.License.SetNonCommercialPersonal("DIMAS RIO SETIAWAN");
        var stream = new MemoryStream();

        using (var package = new ExcelPackage(stream))
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");

            ws.Cells[1, 1].Value = "No";
            ws.Cells[1, 2].Value = "Nama";

            var data = new List<string> { "Satu", "Dua", "Tiga" };

            for (int i = 0; i < data.Count; i++)
            {
                ws.Cells[i + 2, 1].Value = i + 1;
                ws.Cells[i + 2, 2].Value = data[i];
            }

            package.Save();
        }

        stream.Position = 0;
        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Contoh-Excel.xlsx");
    }

}
