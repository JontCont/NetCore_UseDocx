using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using UseDocx.Models;
using NPOI.XWPF.UserModel;
using System.Data;

namespace UseDocx.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _env;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment env)
        {
            _logger = logger;
            _env = env;
        }

        public IActionResult Index()
        {
            string docxPath = _env.WebRootPath + "\\upload\\PR.docx";

            if (System.IO.File.Exists(docxPath))
            {
                FileStream fs = new (docxPath, FileMode.Open, FileAccess.Read);
                XWPFDocument docx = new (fs);
                foreach (var bodyItem in docx.BodyElements)
                {
                    switch (bodyItem.ElementType)
                    {
                        case BodyElementType.TABLE:
                            Set_DocxTableText(bodyItem.Body);
                            break;
                        case BodyElementType.PARAGRAPH:
                            Set_DocxText(bodyItem.Body);
                            break;
                        case BodyElementType.CONTENTCONTROL:break;
                        default:break;
                    }
                }

                return Download(docx); 
            }
            return View();
        }

        public void Set_DocxText(IBody docx)
        {
            foreach (var para in docx.Paragraphs)
            {
                string oldtext = para.ParagraphText;
                string newText = "趙錢孫";
                if (oldtext == "")
                    continue;
                string temptext = para.ParagraphText;
                //以下為替換文件模版中的關鍵字
                if (temptext.Contains("[$name$]"))
                    temptext = temptext.Replace("[$name$]", newText);
                para.ReplaceText(oldtext, temptext);
            }
        }

        public void Set_DocxTableText(IBody docx)
        {
            foreach (XWPFTable dt in docx.Tables)
            {
                foreach (XWPFTableRow dr in dt.Rows)
                {
                    foreach (XWPFTableCell dc in dr.GetTableICells())
                    {
                        foreach (var para in dc.Paragraphs)
                        {
                            string oldtext = para.ParagraphText;
                            string newText = "趙錢孫";
                            if (oldtext == "")
                                continue;
                            string temptext = para.ParagraphText;
                            //以下為替換文件模版中的關鍵字
                            if (temptext.Contains("[$name$]"))
                                temptext = temptext.Replace("[$name$]", newText);
                            para.ReplaceText(oldtext, temptext);
                        }
                    }
                }
            }
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


        public async Task<IActionResult> DownloadAsync(string filePath)
        {
            var memoryStream = new MemoryStream();
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                await stream.CopyToAsync(memoryStream);
            }
            memoryStream.Seek(0, SeekOrigin.Begin);

            // 回傳檔案到 Client 需要附上 Content Type，否則瀏覽器會解析失敗。
            return new FileStreamResult(memoryStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        }

        public IActionResult Download(XWPFDocument fs)
        {
            var memoryStream = new MemoryStream();
            fs.Write(memoryStream);
            memoryStream.Seek(0, SeekOrigin.Begin);
            // 回傳檔案到 Client 需要附上 Content Type，否則瀏覽器會解析失敗。
            return new FileStreamResult(memoryStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        }
    }
}
