using System;
using System.Collections.Generic;
using System.IO;

namespace CSharpExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var personList = new List<Person>
            {
                new Person("Ksenia", "Данилкина", 28, "79612213455"),
                new Person("Karina", "Ivanova", 29, "79682213455"),
                new Person("Ivan", "Ivanov", 55, "79682217855"),
                new Person("Oleg", "Zykov", 49, "79786213455")
            };

            var table = new ExcelGenerator().Generate(personList);

            try
            {
                File.WriteAllBytes("Students.xlsx", table);
            }
            catch (IOException)
            {
                Console.WriteLine("Возникла ошибка при выводе данных в файл.");
            }
        }
    }
}
