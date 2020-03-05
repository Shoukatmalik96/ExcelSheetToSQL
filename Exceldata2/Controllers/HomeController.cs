using Exceldata2.CommonCode;
using Exceldata2.Models;
using Exceldata2.Models.Services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Exceldata2.Models.ExcelSheetModel;

namespace Exceldata2.Controllers
{
	public class HomeController : Controller
	{
		public ActionResult Index()
		{
			return View();
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
        [HttpGet]
		public ActionResult Upload()
		{
			ViewBag.Message = "Your contact page.";

			return View();
		}
        [HttpPost]
        public ActionResult Upload(FormCollection formCollection)
        {
            if (Request != null)
            {
                HttpPostedFileBase file = Request.Files["UploadedFile"];
      
                //check  sheet column and date format.
                bool isArabicDateFormat = (formCollection["arabicDate"] != null && formCollection["arabicDate"] == "on");
                bool isTwoColumnSheet = (formCollection["twoColumnFormat"] != null && formCollection["twoColumnFormat"] == "on");
                string year = formCollection["year"];
                if (file != null && !string.IsNullOrEmpty(file.FileName))
                {
                    string Message = string.Empty;
                    string fileName = file.FileName;
                    string fileContentType = file.ContentType;
                    byte[] fileBytes = new byte[file.ContentLength];
                    var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));
                    using (var package = new ExcelPackage(file.InputStream))
                    {
                      var currentSheet = package.Workbook.Worksheets;
                      var TotalExcelSheets = currentSheet.ToList();
                      ExcelSheetModel excel = new ExcelSheetModel();
                        foreach (var item in TotalExcelSheets)
                        {
                            var noOfCol = item.Dimension.End.Column;
                            var noOfRow = item.Dimension.End.Row;

                            if (isTwoColumnSheet)
                            {
                                // this is for File 2019 
                                //int ctr = 1; 
                                // this is for File 2019 
                                int ctr = 2;

                                for (int i = ctr; i <= noOfRow; i++)
                                {
                                    ExcelSheet excelSheet = new ExcelSheet();
                                    if (isArabicDateFormat) {
                                        var date = DateTimeHelper.GetArabicDate(item.Name, year);
                                        excelSheet.Date = DateTime.ParseExact(date, "d-M-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                    }
                                    else
                                    {
                                        var day   = DateTimeHelper.GetDayMonthFromCurrentDate(item.Name, true, false);
                                        var month = DateTimeHelper.GetDayMonthFromCurrentDate(item.Name, false, true);

                                        if (day.Length < 2 && month.Length < 2) {
                                            excelSheet.Date = DateTime.ParseExact(item.Name, "d-M-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                        }
                                        else if (day.Length == 2 && month.Length < 2) {
                                            excelSheet.Date = DateTime.ParseExact(item.Name, "dd-M-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                        }
                                        else if (day.Length < 2 && month.Length == 2) {
                                            excelSheet.Date = DateTime.ParseExact(item.Name, "d-MM-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                        }
                                        else {
                                            excelSheet.Date = DateTime.ParseExact(item.Name, "dd-MM-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                        }
                                    }
                                  // excel sheet rows and columns
                                  var col1 = item.Cells[i, 1].Value;
                                  var col2 = item.Cells[i, 2].Value;
                                  excelSheet.Company = col1 != null ? col1.ToString() : "";
                                  excelSheet.CurrentPercentage = col2 != null ? col2.ToString().Replace("%", "") : "";
                                  excel.ExcelDataList.Add(excelSheet);
                                  ctr++;
                                }
                            }
                            else
                            {
                                int ctr = 3;
                                for (int i = ctr; i <= noOfRow; i++)
                                {
                                    ExcelSheet excelSheet = new ExcelSheet();
                                    excelSheet.Date = DateTime.ParseExact(item.Name, "dd-MM-yyyy", System.Globalization.CultureInfo.InvariantCulture);

                                    var day   = DateTimeHelper.GetDayMonthFromCurrentDate(item.Name, true, false);
                                    var month = DateTimeHelper.GetDayMonthFromCurrentDate(item.Name, false, true);

                                    if (day.Length < 2 && month.Length < 2){
                                        excelSheet.Date = DateTime.ParseExact(item.Name, "d-M-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                    }
                                    else if (day.Length == 2 && month.Length < 2){
                                        excelSheet.Date = DateTime.ParseExact(item.Name, "dd-M-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                    }
                                    else if (day.Length < 2 && month.Length == 2){
                                        excelSheet.Date = DateTime.ParseExact(item.Name, "d-MM-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                    }
                                    else{
                                        excelSheet.Date = DateTime.ParseExact(item.Name, "dd-MM-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                    }
                                  var col1 = item.Cells[i, 1].Value;
                                  var col2 = item.Cells[i, 2].Value;
                                  var col3 = item.Cells[i, 4].Value;
                                  excelSheet.StockSymbol = col1 != null ? col1.ToString() : "";
                                  excelSheet.Company = col2 != null ? col2.ToString() : "";
                                  excelSheet.CurrentPercentage = col3 != null ? col3.ToString().Replace("%", "") : "";
                                  excel.ExcelDataList.Add(excelSheet);
                                  ctr++;
                                }
                            }
                        }
                        // complete excel work book data
                        var list = excel.ExcelDataList.ToList();
                        try
                        {
                            if (list != null) {
                                // Adding excel data to database
                                ExcelSheetsService.Instance.AddExcelSheet(list);
                                //var isDataExist = ExcelService.Instance.GetAllExcelSheetData();
                                //if (isDataExist.Count() > 0)
                                //{
                                //    Isempty = ExcelService.Instance.DeleteExcelSheetData();
                                //}
                                //if (Isempty)
                                //{
                                //    ExcelService.Instance.AddExcelSheet(list);
                                //}V
                            }
                            else
                            {
                                Message = "Model is Empty !";
                            }
                        }
                        catch (Exception ex)
                        {
                            Message = ex.Message;
                        }
                    }
                }
            
            }
            return View("Index");
        }
         public Boolean IsNumber(String value)
    {
        return value.All(Char.IsDigit);
    }
    }
   
}

