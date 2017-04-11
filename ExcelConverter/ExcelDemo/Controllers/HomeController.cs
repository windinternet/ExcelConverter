using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.IO;
using ExcelConverter;
namespace ExcelDemo.Controllers
{
    public class HomeController : Controller
    {
        //
        // GET: /Home/

        public ActionResult Index()
        {
            return View();
        }
        public ActionResult PriviewExcel()
        {
            //导入
            if (Request.Files.Count==0)
            {
                return Redirect("/");
            }
            var file = Request.Files[0];
            var fileName=Guid.NewGuid().ToString()+Path.GetExtension(file.FileName);
            var realPath=Server.MapPath(fileName);
            if (!Directory.Exists(Path.GetDirectoryName(realPath)))
                Directory.CreateDirectory(Path.GetDirectoryName(realPath));
            file.SaveAs(realPath);
            ExcelDocument<Models.FileInfo> fileinfos = ExcelDocument<Models.FileInfo>.LoadFromFile<Models.FileInfo>(realPath);
            ViewData.Add("infos",fileinfos);
            if (System.IO.File.Exists(realPath))
            {
                System.IO.File.Delete(realPath);
            }
                return View();
        }
        public void ExportExcel()
        {
            //导出
            var path = Request.Params["path"];
            var fileName = "Excel导出Demo";
            var files = Directory.GetFiles(path);
            ExcelDocument<FileInfo> fileinfos = new ExcelDocument<FileInfo>("目录文件列表");
            foreach (var file in files)
            {
                fileinfos.WriteRow(new FileInfo(file));
            }
            Response.ClearHeaders();
            Response.Clear();
            Response.AppendHeader("content-disposition", "attachment; filename=" + (fileName + ".xls"));
            Response.ContentType = "application/vnd.ms-excel";
            Response.ContentEncoding = Encoding.UTF8;
            Response.Write("<html>\r\n");
            Response.Write("<head>\r\n");
            Response.Write("<meta http-equiv=\"content-type\" content=\"application/ms-excel; charset=UTF-8\"/>\r\n");
            Response.Write("</head>\r\n");
            Response.Write("<body>\r\n");
            Response.Write(fileinfos.ToHtml());
            Response.Write("</body>\r\n");
            Response.Write("</html>\r\n");
            Response.End();
        }
    }
}
