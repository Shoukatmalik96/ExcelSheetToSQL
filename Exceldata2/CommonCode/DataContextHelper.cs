//using ExcelData.DataServcies;
using ExcelData.DataServcies;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Exceldata2.CommonCode
{
	public class DataContextHelper
	{

        public static ExcelFileConnectionStringDB GetPPDataContext(bool enableAutoSelect = false)
        {
            return (GetNewDataContext("ExcelFileConnectionString", enableAutoSelect));
        }

        private static ExcelFileConnectionStringDB GetNewDataContext(string connectionStringName, bool enableAutoSelect)
        {
            ExcelFileConnectionStringDB repository = new ExcelFileConnectionStringDB(connectionStringName);
            repository.EnableAutoSelect = enableAutoSelect;
            return (repository);
        }
    }
}