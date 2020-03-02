using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Exceldata2.Models
{
	public class ExcelModel
	{

		public ExcelSheet Excel { get; set; }

		public List<ExcelSheet> ExcelDataList { get; set; } = new List<ExcelSheet>();
		public List<ExcelSheet> excelList { get; set; } = new List<ExcelSheet>();
	}

	public class ExcelSheet
	{
		public DateTime Date { get; set; }
		public string  CompanyCode { get; set; }
		public string  Company { get; set; }
		public string HighestRate { get; set; }
		public string CurrentOwnership { get; set; }
		public string ForeignStrategicInvestorOwnership { get; set; }
	}
}