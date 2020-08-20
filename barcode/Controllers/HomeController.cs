using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;

using barcode.Models;

using Common;

namespace barcode.Controllers
{
    public class HomeController : Controller
    {
        [HttpPost]
        public IActionResult Index(List<IFormFile> BarcodeSourceFile)
        {
            if (BarcodeSourceFile.Count == 1)
            {
                var inputSream = BarcodeSourceFile[0].OpenReadStream();
                MemoryStream outputStream = new MemoryStream();
                BarcodeXlsxImporter barcodeXlsx = new BarcodeXlsxImporter();
                barcodeXlsx.Convert(inputSream, outputStream);
                outputStream.Seek(0, SeekOrigin.Begin);
                return File(outputStream, "application/excel", DateTime.Now.ToString("yyyyMMddhhmmss") +  ".xlsx");
            }

            return View();
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Convert()
        {
            return View();
        }

        public IActionResult HowTo()
        {
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
    }
}
