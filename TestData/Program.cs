﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelData;

namespace TestData
{
    class Program
    {
        static void Main(string[] args)
        {
            var pr = ExcelBook.ExcelBookTest();
            Console.WriteLine(pr);
            Console.ReadKey();
        }
    }
}