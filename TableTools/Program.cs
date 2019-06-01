using OfficeOpenXml;
using System;
using System.Threading;
namespace TableTools
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            TableRead.Init();
            while (true) {
                ReadLine();
                Thread.Sleep(1);
            }
        }


        static void ReadLine() {


        }
    }
}
