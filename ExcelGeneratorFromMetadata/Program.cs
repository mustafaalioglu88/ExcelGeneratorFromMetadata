using System;
using System.Collections.Generic;

namespace ExcelGeneratorFromMetadata
{
    internal class Program
    {
        private static CrmHelper helper;

        private static void Main(string[] args)
        {
            var version = " v." + System.Reflection.Assembly.GetExecutingAssembly()
                                           .GetName()
                                           .Version;
            Console.WriteLine("Executable version is {0}", version);
            var fileLocation = AppDomain.CurrentDomain.BaseDirectory + "list.txt";
            var entityLogicalNameList = FileHelper.GetListFromFile(fileLocation);
            if (entityLogicalNameList.Count == 0)
            {
                Console.WriteLine("There is no entity to model in list. Please put a 'list.txt' to root of executable\n" +
                                  "And make sure you have a folder named ");
            }
            else
            {
                Console.Write("Please type organization url: ");
                var url = Console.ReadLine();
                Console.Write("Please type organization user domain: ");
                var domain = Console.ReadLine();
                Console.Write("Please type organization username: ");
                var userName = Console.ReadLine();
                Console.Write("Please type organization user password: ");
                var password = Console.ReadLine();
                if (url == domain && userName == password)
                {
                    GetDefaultConfiguration(ref url, ref domain, ref userName, ref password);
                }

                IList<string> skippedList;
                helper = new CrmHelper(url, domain, userName, password);
                var list = helper.GetEntityModel(entityLogicalNameList, out skippedList);
                FileHelper.WriteToExcel(list);
            }

            Console.WriteLine("Done. Press a key to exit.");
            Console.ReadLine();
        }

        private static void GetDefaultConfiguration(ref string url, ref string domain, ref string userName,
            ref string password)
        {
            url = "";
            domain = "";
            userName = "";
            password = "";
        }
    }
}