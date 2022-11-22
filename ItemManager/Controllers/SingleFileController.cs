using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ItemManager.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using DataSolutions.Logging.Logger;

namespace ItemManager.Controllers
{
    public class SingleFileController : Controller
    {
        private readonly IHostingEnvironment _env;

        public SingleFileController(IHostingEnvironment env)
        {
            _env = env;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Index(LacosteFile item)
        {

            if (ModelState.IsValid)
            {
                IFormFile file_for_processing = item.File;
                //Check file extension is a photo
                string extension =
                       Path.GetExtension(file_for_processing.FileName);

                if (file_for_processing.Length > 0) //ensure the file is not empty
                {
                    string filePath = Path.Combine(_env.ContentRootPath, "TEST", "ftp", "Upload"
                                                , file_for_processing.FileName);

                    //write file to file system
                    using (FileStream fs = new FileStream(filePath, FileMode.Create))
                    {
                        await file_for_processing.CopyToAsync(fs);
                    }

                    //Combining
                    var payload_LACOSTEPostProcess = new LACOSTEPostprocessCycle(_env);
                    try
                    {
                        payload_LACOSTEPostProcess.DoWork();
                        TempData["MsgChangeStatus"] += "Records have been successfully inserted into DB";
                        return View("Index");
                    }
                    catch (Exception ex)
                    {
                        TempData["MsgChangeStatus"] += ex.ToString();
                        return View("Index");
                        throw;
                    }
                }
                return View();
            }
            return View();
        }

    }
}