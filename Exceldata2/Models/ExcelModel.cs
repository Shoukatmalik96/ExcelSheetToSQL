using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Exceldata2.Models
{
	public class ExcelSheetModel
	{

		public ExcelSheet Excel { get; set; }
		public List<ExcelSheet> ExcelDataList { get; set; } = new List<ExcelSheet>();
		public List<ExcelSheet> excelList { get; set; } = new List<ExcelSheet>();
	}

	public class ExcelSheet
	{
		public DateTime Date { get; set; }
		public string StockSymbol { get; set; }
		public string  Company { get; set; }
		public string CurrentPercentage { get; set; }
	}

	public class ExcelViewModel
	{
		public string twoColumnFormat { get; set; }
		public string arabicDate { get; set; }

		public string year { get; set; }
	}
}