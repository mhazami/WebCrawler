using Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;

namespace WebCrawler.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            WebClient webClient = new WebClient();
            string page = webClient.DownloadString("https://www.coingecko.com/tr/borsalar?page=1");

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(page);

            List<List<string>> table = doc.DocumentNode.SelectSingleNode("//table[@class='sort table mb-0 text-sm text-lg-normal table-scrollable']")
                        .Descendants("tr")
                        .Skip(1)
                        .Where(tr => tr.Elements("td").Count() > 1)
                        .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToList())
                        .ToList();

            GenerateExcel(table);
            return View();
        }

        public static void GenerateExcel(List<List<string>> list)
        {
            try
            {
                var file = new FileInfo("test.xlsx");
                string currentFileName = Path.GetFileName("test.xlsx");
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage excel = new ExcelPackage(file);
                var sheetcreate = excel.Workbook.Worksheets.Add("Sheet1");

                char c = 'A';
                c = (char)(((int)c) + 9);
                //sheetcreate.Cells["A1:D1"].Value = rptHeader;
                sheetcreate.Cells["A1:" + c + "1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheetcreate.Cells[1, 1, 1, 9].Merge = true;
                sheetcreate.Cells[1, 1, 1, 9].Style.Font.Bold = true;

                for (int i = 0; i < list.Count; i++)
                {
                    var row = list[i];
                    i++;

                    for (int j = 1; j < row.Count; j++)
                    {
                        sheetcreate.Cells[i, ++j].Value = row[j];
                        sheetcreate.Cells[i, j].Style.Font.Bold = true;
                        sheetcreate.Cells[i, j].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        sheetcreate.Cells[i, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                }




                sheetcreate.Cells.AutoFitColumns();
                excel.Save();
            }
            catch (Exception e)
            {
            }
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