using Exceldata2.CommonCode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Exceldata2.Models.Services
{
	public class ExcelService
	{

        #region Define as Singleton
        private static ExcelService _Instance;
        public static ExcelService Instance
        {
            get
            {
                if (_Instance == null)
                {
                    _Instance = new ExcelService();
                }

                return (_Instance);
            }
        }
        private ExcelService()
        {
        }
        #endregion

        public void AddExcelSheet(List<ExcelSheet> excelSheets)
        {
            var NewSheet = excelSheets.OrderBy(x => x.Date);
            foreach (var sheet in NewSheet)
            {
                ExcelModel excel = new ExcelModel();
                excel.Excel = sheet;

                using (var context = DataContextHelper.GetPPDataContext())
                {   
                    context.Insert(excel.Excel);
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