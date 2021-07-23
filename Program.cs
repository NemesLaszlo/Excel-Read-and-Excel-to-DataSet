using System;
using System.Data;

namespace ExcelReadTest
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelRead.readXLS(@"C:\Users\Laci\Desktop\Test.xlsx");


            Console.WriteLine("--------------------------------------------------");


            ExcelRead.getExcelFile(@"C:\Users\Laci\Desktop\Test.xlsx");


            Console.WriteLine("\n--------------------------------------------------");



            DataSet DataSet = ExcelRead.DatasetImportFromExcel(@"C:\Users\Laci\Desktop\Test.xlsx");

            Console.WriteLine("Table counts: " + DataSet.Tables.Count);

            // Example to check and get the Table name into a variable
            if (DataSet.Tables.Contains("Sheet1"))
            {
                string name = DataSet.Tables["Munka1"].TableName;
                Console.WriteLine("First Sheet name: " + name);
            }

            foreach (DataRow row in DataSet.Tables["Sheet1"].Rows) // DataSet.Tables[0].Rows is good as well
            {
                object[] rowData = row.ItemArray;
                foreach (var i in rowData)
                {
                    Console.WriteLine(i.ToString());
                }
            }

            Console.WriteLine("Other sheet: ");

            foreach (DataRow row in DataSet.Tables["Sheet2"].Rows)
            {
                object[] rowData = row.ItemArray;
                foreach (var i in rowData)
                {
                    Console.WriteLine(i.ToString());
                }
            }

            Console.ReadKey();
        }
    }
}
