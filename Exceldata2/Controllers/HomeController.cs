using Exceldata2.Models;
using Exceldata2.Models.Services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Exceldata2.Models.ExcelModel;

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

                if (file != null && !string.IsNullOrEmpty(file.FileName))
                {
                    string fileName = file.FileName;
                    string fileContentType = file.ContentType;
                    byte[] fileBytes = new byte[file.ContentLength];
                    var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));

                    using (var package = new ExcelPackage(file.InputStream))
                    {
                       
                        var currentSheet = package.Workbook.Worksheets;
                        var TotalExcelSheets = currentSheet.ToList();

                        ExcelModel excel = new ExcelModel();

                        
                        foreach (var item in TotalExcelSheets)
                        {
                        
                            int ctr = 3;

                            var noOfCol = item.Dimension.End.Column;
                            var noOfRow = item.Dimension.End.Row;

                            //ExcelSheet excelSheet = new ExcelSheet();
                            
                            
                            for (int i = ctr ; i <= noOfRow; i++)
                            {

                                ExcelSheet excelSheet = new ExcelSheet();
                                excelSheet.Date = DateTime.ParseExact(item.Name, "dd-MM-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                var col1 = item.Cells[i, 1].Value;
                                var col2 = item.Cells[i, 2].Value;
                                var col3 = item.Cells[i, 3].Value;
                                var col4 = item.Cells[i, 4].Value;
                                var col5 = item.Cells[i, 5].Value;
                                excelSheet.CompanyCode = col1 != null ? col1.ToString() : "";
                                excelSheet.Company = col2 != null ? col2.ToString() : "";
                                excelSheet.CurrentOwnership = col3 != null ? col3.ToString() : "";
                                excelSheet.HighestRate = col4 != null ? col4.ToString() : "";
                                excelSheet.ForeignStrategicInvestorOwnership = col5 != null ? col5.ToString() : "";

                                excel.ExcelDataList.Add(excelSheet);
                                ctr++;
                            }
                           
                        }
                        var list = excel.ExcelDataList.ToList();
                        string Message = string.Empty;
                        try
                        {
                           
                            if (list != null)
                            {
                                ExcelService.Instance.AddExcelSheet(list);
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
    }
}