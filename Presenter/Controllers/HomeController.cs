using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Web.Mvc;
using OFFICECORE = Microsoft.Office.Core;
using POWERPOINT = Microsoft.Office.Interop.PowerPoint;


namespace Presenter.Controllers
{
    [RequireHttps]
    public class HomeController : Controller
    {
        public void convertPPT(String FilePath, String NewFolderPath)
        {
            POWERPOINT.Application App = new Microsoft.Office.Interop.PowerPoint.Application();
            POWERPOINT.Presentation pres = App.Presentations.Open(FilePath, OFFICECORE.MsoTriState.msoTrue, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse);
            pres.SaveAs(NewFolderPath, POWERPOINT.PpSaveAsFileType.ppSaveAsJPG, OFFICECORE.MsoTriState.msoFalse);
            pres.Close();

        }
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Upload()
        {
            ViewBag.Message = "Please upload your file.";

            return View();
        }


       
        [HttpPost]
        public ActionResult UploadFile(HttpPostedFileBase file)
        {
            
            if (file != null && file.ContentLength > 0)
            {
                var fileName = Path.GetFileName(file.FileName);
                var path = Path.Combine(Server.MapPath("~/App_Data/uploads"), fileName);
                file.SaveAs(path);
                convertPPT(path, Server.MapPath("~/App_Data/uploads/1"));
            }

            return RedirectToAction("Upload");
        }
   


        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}