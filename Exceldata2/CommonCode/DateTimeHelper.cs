using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Exceldata2.CommonCode
{
	public static class DateTimeHelper
	{

        public static string GetDayMonthFromCurrentDate(string date, bool day, bool month)
        {
            string[] Words = date.Split(new char[] { '-' });
            if (day == true)
            {
                return Words[0];
            }

            return Words[1];
        }

        public static string GetArabicDate(string dt, string year)
        {
            //var day = dt.Split(' ')[0];
            var day = dt.Trim().Split(' ')[0];
            string month = TransformArabicMonthToEnglish(dt.Trim().Split(' ')[1]);
            string englishDate = string.Format("{0}-{1}-{2}", day, month, year);
            return englishDate;
        }

        public static string TransformArabicMonthToEnglish(string monthAr)
        {
            string monthEn = null;

            if (monthAr == "يناير")
            {
                monthEn = "1";
            }
            else if (monthAr == "فبراير")
            {
                monthEn = "2";
            }
            else if (monthAr == "مارس")
            {
                monthEn = "3";
            }
            else if (monthAr == "أبريل" || monthAr == "ابريل")
            {
                monthEn = "4";
            }
            else if (monthAr == "مايو")
            {
                monthEn = "5";
            }
            else if (monthAr == "يونيو")
            {
                monthEn = "6";
            }
            else if (monthAr == "يوليو")
            {
                monthEn = "7";
            }
            else if (monthAr == "أغسطس" || monthAr == "اغسطس")
            {
                monthEn = "8";
            }
            else if (monthAr == "سبتمبر")
            {
                monthEn = "9";
            }
            else if (monthAr == "اكتوبر")
            {
                monthEn = "10";
            }
            else if (monthAr == "نوفمبر")
            {
                monthEn = "11";
            }
            else if (monthAr == "ديسمبر")
            {
                monthEn = "12";
            }
            return monthEn;
        }
    }
}