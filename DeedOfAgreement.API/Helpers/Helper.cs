using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace DeedOfAgreement.API.Helpers
{
    public static class Helper
    {
        private static readonly string ConnectionString = ConfigurationManager.ConnectionStrings["ERPDbContext"].ConnectionString;
        public static DataSet GetDataFromBackDB(string sql)
        {
            var dataset = new DataSet();
            using (var sqlConnection = new SqlConnection(ConnectionString))
            {
                sqlConnection.Open();
                var sqlCommand = new SqlCommand(sql, sqlConnection);
                var adapter = new SqlDataAdapter(sqlCommand);
                adapter.Fill(dataset);
                sqlCommand.Dispose();
                sqlConnection.Dispose();
                sqlConnection.Close();
            }
            return dataset;
        }
        private static string NumberToWords(int number)
        {
            if (number == 0)
                return "Zero";

            if (number < 0)
                return "Minus " + NumberToWords(Math.Abs(number));

            string words = "";

            if ((number / 1000) > 0)
            {
                words += NumberToWords(number / 1000) + " Thousand ";
                number %= 1000;
            }

            if ((number / 100) > 0)
            {
                words += NumberToWords(number / 100) + " Hundred ";
                number %= 100;
            }

            if (number > 0)
            {
                string[] unitsMap = { "Zero", "One", "Two", "Three", "Four", "Five", "Six",
                                  "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve",
                                  "Thirteen", "Fourteen", "Fifteen", "Sixteen",
                                  "Seventeen", "Eighteen", "Nineteen" };

                string[] tensMap = { "Zero", "Ten", "Twenty", "Thirty", "Forty",
                                 "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };

                if (number < 20)
                    words += unitsMap[number];
                else
                {
                    words += tensMap[number / 10];
                    if ((number % 10) > 0)
                        words += " " + unitsMap[number % 10];
                }
            }

            return words.Trim();
        }
        private static string GetOrdinal(int number)
        {
            if (number % 100 >= 11 && number % 100 <= 13)
                return number + "th";

            switch (number % 10)
            {
                case 1: return number + "st";
                case 2: return number + "nd";
                case 3: return number + "rd";
                default: return number + "th";
            }
        }
        public static string GetDateInFullWords(DateTime date)
        {
            string day = GetOrdinal(date.Day);
            string month = date.ToString("MMMM");
            string year = NumberToWords(date.Year);

            return $"{day} Day of {month} {year}";
        }
        public static List<T> ToList<T>(this DataTable data) where T : class
        {
            PropertyDescriptorCollection properties =
                TypeDescriptor.GetProperties(typeof(T));
            var list = new List<T>();
            foreach (var row in data.AsEnumerable())
            {
                var item = Activator.CreateInstance(typeof(T));
                foreach (PropertyDescriptor property in properties)
                {
                    try
                    {
                        property.SetValue(item, row[property.Name]);
                    }
                    catch (Exception ex)
                    {

                    }
                }
                list.Add((T)item);
            }
            return list;
        }
    }
}