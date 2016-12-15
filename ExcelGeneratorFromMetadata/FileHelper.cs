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
                foreach (var entityModel in list)
                {
                    var excel = new Application();
                    excel.Visible = true;
                    excel.DisplayAlerts = false;
                    var wb = excel.Workbooks.Add();
                    excel.DisplayAlerts = true;

                    Console.WriteLine(entityModel.DisplayName + " is now exporting.");
                    Worksheet sh = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                    var sheetName = CleanUp("Entity");
                    if(sheetName.Length > 31)
                    {
                        sheetName = sheetName.Substring(0, 31);
                    }

                    sh.Name = sheetName;
                    sh.Cells[1, "A"].Value2 = "LogicalName";
                    sh.Cells[1, "A"].Font.Bold = true;
                    sh.Cells[1, "B"].Value2 = "DisplayName";
                    sh.Cells[1, "B"].Font.Bold = true;
                    sh.Cells[1, "C"].Value2 = "Description";
                    sh.Cells[1, "C"].Font.Bold = true;
                    sh.Cells[1, "D"].Value2 = "DisplayNamePlural";
                    sh.Cells[1, "D"].Font.Bold = true;


                    sh.Cells[2, "A"].Value2 = entityModel.LogicalName;
                    sh.Cells[2, "B"].Value2 = entityModel.DisplayName;
                    sh.Cells[2, "C"].Value2 = entityModel.Description;
                    sh.Cells[2, "D"].Value2 = entityModel.DisplayNamePlural;

                    sh.Cells[3, "A"].Value2 = "LogicalName";
                    sh.Cells[3, "A"].Font.Bold = true;
                    sh.Cells[3, "B"].Value2 = "DisplayName";
                    sh.Cells[3, "B"].Font.Bold = true;
                    sh.Cells[3, "C"].Value2 = "Description";
                    sh.Cells[3, "C"].Font.Bold = true;
                    sh.Cells[3, "D"].Value2 = "AttributeType";
                    sh.Cells[3, "D"].Font.Bold = true;
                    sh.Cells[3, "E"].Value2 = "LookupEntityLogicalName";
                    sh.Cells[3, "E"].Font.Bold = true;
                    sh.Cells[3, "F"].Value2 = "OptionSetList";
                    sh.Cells[3, "F"].Font.Bold = true;
                    sh.Cells[3, "G"].Value2 = "GlobalOptionSetListLogicalName";
                    sh.Cells[3, "G"].Font.Bold = true;
                    sh.Cells[3, "H"].Value2 = "MinValue";
                    sh.Cells[3, "H"].Font.Bold = true;
                    sh.Cells[3, "I"].Value2 = "MaxValue";
                    sh.Cells[3, "I"].Font.Bold = true;
                    sh.Cells[3, "J"].Value2 = "IsRequired";
                    sh.Cells[3, "J"].Font.Bold = true;
                    sh.Cells[3, "K"].Value2 = "OtherDisplayName";
                    sh.Cells[3, "K"].Font.Bold = true;
                    sh.Cells[3, "L"].Value2 = "OtherDescription";
                    sh.Cells[3, "L"].Font.Bold = true;

                    var i = 4;
                    foreach (var attributeModel in entityModel.AttributeModelList)
                    {
                        sh.Cells[i, "A"].Value2 = attributeModel.LogicalName;
                        sh.Cells[i, "B"].Value2 = attributeModel.DisplayName;
                        sh.Cells[i, "C"].Value2 = attributeModel.Description;
                        sh.Cells[i, "D"].Value2 = attributeModel.DataType;
                        sh.Cells[i, "E"].Value2 = attributeModel.LookupEntityLogicalName;
                        sh.Cells[i, "F"].Value2 = attributeModel.OptionSetList;
                        sh.Cells[i, "G"].Value2 = attributeModel.GlobalOptionSetListLogicalName;
                        sh.Cells[i, "H"].Value2 = attributeModel.MinValue;
                        sh.Cells[i, "I"].Value2 = attributeModel.MaxValue;
                        sh.Cells[i, "J"].Value2 = attributeModel.IsRequired ? "yes" : "";
                        sh.Cells[i, "K"].Value2 = attributeModel.OtherDisplayName;
                        sh.Cells[i, "L"].Value2 = attributeModel.OtherDescription;
                        i++;
                    }

                    var range = sh.Range["A1", "L" + i];
                    range.Columns.AutoFit();

                    Worksheet sh2 = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                    sh2.Name = "WebResource";

                    excel.DisplayAlerts = false;
                    var worksheetDelete = (Worksheet)wb.Worksheets[1];
                    worksheetDelete.Delete();
                    excel.DisplayAlerts = true;

                    var path = AppDomain.CurrentDomain.BaseDirectory + "Output\\";
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    wb.SaveAs(path + entityModel.LogicalName + ".xlsx");
                    wb.Close(true);
                    excel.Quit();
                }
                
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