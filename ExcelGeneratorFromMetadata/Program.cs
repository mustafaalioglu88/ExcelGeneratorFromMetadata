using System;

namespace ExcelGeneratorFromMetadata
{
    internal class Program
    {
        private static readonly bool includeOnlycustomPrefixedAttributes = true;
        private static readonly string customPrefix = "pfx_";
        private static CrmHelper helper;

        private static void Main(string[] args)
        {
            var fileLocation = AppDomain.CurrentDomain.BaseDirectory + "list.txt";
            var entityLogicalNameList = FileHelper.GetListFromFile(fileLocation);
            if (entityLogicalNameList.Count == 0)
            {
                Console.WriteLine("There is no entity to model in list.");
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
                helper = new CrmHelper(url, domain, userName, password);
                var list = helper.GetEntityModel(entityLogicalNameList, includeOnlycustomPrefixedAttributes,
                    customPrefix);
                FileHelper.WriteToExcel(list);
            }

            Console.WriteLine("Done. Press a key to exit.");
            Console.ReadLine();
        }
    }
}