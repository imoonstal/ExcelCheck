using Business.Handler;
using Excel.Handler;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = Console.ReadLine();
            string result = Checking.CheckExcel(path);
            //string test = ExcelHelper.GetStringByCell(path, "宣传情况", 0, 0);
            //Console.WriteLine(result);
            Console.ReadLine();
        }
    }
}
