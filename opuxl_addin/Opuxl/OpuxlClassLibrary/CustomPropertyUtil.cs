using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpuxlClassLibrary
{
    class CustomPropertyUtil
    {
        public void SetCustomProp(Worksheet worksheet, string key, string value)
        {
            // We always have to delete an existing prop before writing a new one. Otherwise we are getting an exception.
            DeleteCustomProp(worksheet, key);
            worksheet.CustomProperties.Add(key, value);
        }

        public string GetCustomProp(Worksheet worksheet, string key)
        {
            for (int i = 1; i <= worksheet.CustomProperties.Count; i++)
            {
                if (worksheet.CustomProperties.get_Item(i).Name == key)
                    return (worksheet.CustomProperties.get_Item(i).Value);
            }
            return string.Empty;
        }

        public void DeleteCustomProp(Worksheet worksheet, string key)
        {
            for (int i = 1; i <= worksheet.CustomProperties.Count; i++)
            {
                if (worksheet.CustomProperties.Item[i].Name == key)
                {
                    worksheet.CustomProperties.Item[i].Delete();
                    break;
                }
            }
        }
    }
}
