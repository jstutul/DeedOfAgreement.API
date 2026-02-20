using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;

namespace DeedOfAgreement.API.Utility
{
    public static class Helpers
    {
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
                    if (data.Columns.Contains(property.Name))
                    {
                        if (row[property.Name] != DBNull.Value)
                        {
                            property.SetValue(item, row[property.Name]);
                        }
                    }
                }
                list.Add((T)item);
            }
            return list;
        }
    }
}
