using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using barcode.Models;


namespace barcode.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            if (this.Request.Method == "POST")
            {
                foreach (var form in this.Request.Form.Files)
                {
                    Debug.WriteLine("Index, {0}, {1}", form.Key, form.Value);
                }
            }

            return View();
        }

        public IActionResult Convert()
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
