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