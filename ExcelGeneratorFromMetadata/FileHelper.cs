using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ExcelGeneratorFromMetadata
{
    public class FileHelper
    {
        public static List<string> GetListFromFile(string fileLocation)
        {
            var list = new List<string>();
            string line;

            StreamReader file;
            try
            {
                file = new StreamReader(fileLocation);
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("Can not locate " + fileLocation + " file.");
                return new List<string>();
            }

            while ((line = file.ReadLine()) != null)
            {
                line = CleanInput(line);
                if (!string.IsNullOrWhiteSpace(line))
                {
                    list.Add(line);
                }
            }

            file.Close();

            return list;
        }

        public static void WriteToExcel(List<EntityModel> list)
        {
            Console.WriteLine("Excel export started.");
            try
            {
                var excel = new Application();
                excel.Visible = true;
                var wb = excel.Workbooks.Add();
                foreach (var entityModel in list)
                {
                    Console.WriteLine(entityModel.DisplayName + " is now exporting.");
                    Worksheet sh = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                    var sheetName = CleanUp(entityModel.DisplayName);
                    if(sheetName.Length > 31)
                    {
                        sheetName = sheetName.Substring(0, 31);
                    }

                    sh.Name = sheetName;
                    sh.Cells[1, "A"].Value2 = entityModel.LogicalName;
                    sh.Cells[1, "B"].Value2 = entityModel.DisplayName;

                    sh.Cells[2, "A"].Value2 = "Alan Adı";
                    sh.Cells[2, "A"].Font.Bold = true;
                    sh.Cells[2, "B"].Value2 = "Tanımı";
                    sh.Cells[2, "B"].Font.Bold = true;
                    sh.Cells[2, "C"].Value2 = "Data Tipi";
                    sh.Cells[2, "C"].Font.Bold = true;
                    sh.Cells[2, "D"].Value2 = "Kısıtlamalar";
                    sh.Cells[2, "D"].Font.Bold = true;
                    sh.Cells[2, "E"].Value2 = "Zorunlu Mu?";
                    sh.Cells[2, "E"].Font.Bold = true;

                    var i = 3;
                    foreach (var attributeModel in entityModel.AttributeModelList)
                    {
                        sh.Cells[i, "A"].Value2 = attributeModel.LogicalName;
                        sh.Cells[i, "B"].Value2 = attributeModel.DisplayName;
                        sh.Cells[i, "C"].Value2 = attributeModel.DataType;
                        sh.Cells[i, "D"].Value2 = attributeModel.Constraint;
                        sh.Cells[i, "E"].Value2 = attributeModel.IsRequired ? "Evet" : "Hayır";
                        i++;
                    }

                    var range = sh.Range["A1", "E" + i];
                    range.Columns.AutoFit();
                }

                wb.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "EntityList.xlsx");
                wb.Close(true);
                excel.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    string.Format("Excel export failed. Check your excel installation on local machine. Does your licence valid for excel?\nDetails:\n{0}", ex.Message));
            }
        }

        private static string CleanUp(string label)
        {
            char[] replaceChars = { ':', '\\', '/', '?', '*', '[', ']' };
            var replacedString = new string(label.Where(c => !replaceChars.Contains(c)).ToArray());
            return replacedString;
        }

        private static string CleanInput(string str)
        {
            var sb = new StringBuilder();
            foreach (var c in str)
            {
                if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '.' || c == '_')
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }
    }
}