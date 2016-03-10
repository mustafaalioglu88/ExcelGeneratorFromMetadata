using System;
using System.Collections.Generic;
using System.IO;
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
                    Worksheet sh = wb.Sheets.Add();
                    sh.Name = entityModel.DisplayName;
                    sh.Cells[1, "A"].Value2 = entityModel.LogicalName;
                    sh.Cells[1, "B"].Value2 = entityModel.DisplayName;
                    var i = 2;
                    foreach (var attributeModel in entityModel.AttributeModelList)
                    {
                        sh.Cells[i, "A"].Value2 = attributeModel.LogicalName;
                        sh.Cells[i, "B"].Value2 = attributeModel.DisplayName;
                        sh.Cells[i, "C"].Value2 = attributeModel.DataType;
                        sh.Cells[i, "D"].Value2 = attributeModel.Constraint;
                        sh.Cells[i, "E"].Value2 = attributeModel.IsRequired ? "Evet" : "Hayır";
                        i++;
                    }
                }

                wb.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "EntityList.xlsx");
                wb.Close(true);
                excel.Quit();
            }
            catch (Exception)
            {
                Console.WriteLine(
                    "Excel export failed. Check your excel installation on local machine. Does your licence valid for excel?");
            }
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