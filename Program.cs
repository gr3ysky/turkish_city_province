using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.IO;
using System.Text;

namespace main
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Started");
            var dic = new Dictionary<string, List<string>>();
            using (var package = new ExcelPackage(new FileInfo(Path.Combine(Environment.CurrentDirectory, "pk_2018_08_31.xlsx"))))
            {
                var sheet = package.Workbook.Worksheets[1];
                for (var i = 2; i < 73638; i++)
                {
                    var il = sheet.Cells[i, 1].Value.ToString().ToLower(new System.Globalization.CultureInfo("tr")).Trim();
                    if (!dic.ContainsKey(il))
                    {
                        dic.Add(il, new List<string>());
                    }

                    var ilce = sheet.Cells[i, 2].Value.ToString().ToLower(new System.Globalization.CultureInfo("tr")).Trim();

                    var list = dic[il];
                    if (!list.Contains(ilce))
                    {
                        list.Add(ilce);
                    }
                }
            }

            var json = JsonConvert.SerializeObject(dic);

#pragma warning disable SCS0018 // Path traversal: injection possible in {1} argument passed to '{0}'
            File.WriteAllBytes(Path.Combine(Environment.CurrentDirectory,"data.json"),Encoding.UTF8.GetBytes(json));
#pragma warning restore SCS0018 // Path traversal: injection possible in {1} argument passed to '{0}'

        }

    }
}
