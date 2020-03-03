using Exceldata2.CommonCode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Exceldata2.Models.Services
{
	public class ExcelSheetsService
	{

        #region Define as Singleton
        private static ExcelSheetsService _Instance;
        public static ExcelSheetsService Instance
        {
            get
            {
                if (_Instance == null)
                {
                    _Instance = new ExcelSheetsService();
                }

                return (_Instance);
            }
        }
        private ExcelSheetsService()
        {
        }
        #endregion

        public void AddExcelSheet(List<ExcelSheet> excelSheets)
        {
            var NewSheet = excelSheets.OrderBy(x => x.Date);
            foreach (var sheet in NewSheet)
            {
                ExcelSheetModel excelSheet = new ExcelSheetModel();
                excelSheet.Excel = sheet;

                using (var context = DataContextHelper.GetPPDataContext())
                {   
                    context.Insert(excelSheet.Excel);
                }
            }
             
        }

        public List<ExcelSheet> GetAllExcelSheetData()
        {
                List<ExcelSheet> result = null;

                using (var context = DataContextHelper.GetPPDataContext())
                {
                    var sql = PetaPoco.Sql.Builder
                    .Select("*")
                    .From("ExcelSheet");
                    result = context.Fetch<ExcelSheet>(sql).ToList();
                }

            return result;
        }
        public bool DeleteExcelSheetData()
        {
            bool result = false;

            using (var context = DataContextHelper.GetPPDataContext())
            {
            
               context.Execute(@"Delete from  ExcelSheet");
                
            }

            return true;
        }
    }
}