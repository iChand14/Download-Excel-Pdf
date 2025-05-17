using Microsoft.AspNetCore.Mvc;
using ClosedXML.Excel;
using Rotativa.AspNetCore;
using System.IO;
using ExportDemoApp.Models;

namespace ExportDemoApp.Controllers
{
    public class ExportController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult DownloadPDF()
        {
            var model = new LaporanViewModel
            {
                Title = "Laporan PDF",
                Date = DateTime.Now
            };
            return new ViewAsPdf("PdfView", model)
            {
                FileName = "Laporan.pdf",
                PageSize = Rotativa.AspNetCore.Options.Size.A4
            };
        }

        public IActionResult DownloadExcel()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Data");
            worksheet.Cell(1, 1).Value = "Nama";
            worksheet.Cell(1, 2).Value = "Tanggal";

            worksheet.Cell(2, 1).Value = "Laporan Excel";
            worksheet.Cell(2, 2).Value = DateTime.Now.ToString("dd/MM/yyyy");

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return File(stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Laporan.xlsx");
        }
    }
}
