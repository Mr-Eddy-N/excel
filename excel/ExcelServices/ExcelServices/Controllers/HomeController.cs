using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Web.Mvc;
using System.Reflection;
using DocumentFormat.OpenXml;
using ClosedXML.Excel;
using ExcelServices.Models;

namespace ExcelServices.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult Index()
        {
            
            string filePath =Path.Combine(Server.MapPath("~/MyFiles/ex.xlsx"));
            try {
                var workbook = new XLWorkbook();
                var ws = workbook.Worksheets.Add("Column Cells");

                var columnFromWorksheet = ws.Column(1);
                columnFromWorksheet.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
                columnFromWorksheet.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
                columnFromWorksheet.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
                columnFromWorksheet.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

                var columnFromRange = ws.Range("B1:B9").FirstColumn();

                columnFromRange.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
                columnFromRange.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
                columnFromRange.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
                columnFromRange.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

                workbook.SaveAs(filePath);
           
                }
            catch(Exception ex){

                }
            return View();
        }
        [HttpGet]
        public virtual ActionResult DownloadExcel()
        {
            
            string filePath = Path.Combine(Server.MapPath("~/MyFiles/ex.xlsx"));
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Column Cells");

            var columnFromWorksheet = ws.Column(1);
            columnFromWorksheet.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
            columnFromWorksheet.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
            columnFromWorksheet.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
            columnFromWorksheet.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

            var columnFromRange = ws.Range("B1:B9").FirstColumn();

            columnFromRange.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
            columnFromRange.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
            columnFromRange.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
            columnFromRange.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

            workbook.SaveAs(filePath);
            string file ="ex.xlsx";
            string fullPath = Path.Combine(Server.MapPath("~/MyFiles"), file);
            return File(fullPath, "application/vnd.ms-excel", file);
            //return File(workbook,"application/vnd.ms-excel","reportesotee.xlsx");
        }

    }
}