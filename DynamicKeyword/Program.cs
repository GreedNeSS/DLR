using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DynamicKeyword
{
    class Person
    {
        public string FirstName { get; set; } = "";
        public string LastName { get; set; } = "";
    }
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("****** Dynamic ******");

            PrintThreeStrings();
            Console.WriteLine();

            ChangeDynamicDataType();
            Console.WriteLine();

            InvokeMemberOnDynamicData();
            Console.WriteLine();
        }

        static void ImplicitlyTypedVariable()
        {
            var a = new List<int> { 90, 20 };
           
            //a = "Hello";
        }

        static void UseObjectVariable()
        {
            object o = new Person { FirstName = "Mike", LastName = "Larson" };

            Console.WriteLine($"FistName: {((Person)o).FirstName}");
        }

        static void PrintThreeStrings()
        {
            var s1 = "Hello";
            object s2 = "wwww";
            dynamic s3 = "Http";

            Console.WriteLine($"s1.GetType(): {s1.GetType()}");
            Console.WriteLine($"s2.GetType(): {s2.GetType()}");
            Console.WriteLine($"s3.GetType(): {s3.GetType()}");
        }

        static void ChangeDynamicDataType()
        {
            dynamic t = "hello";
            Console.WriteLine($"t.GetType(): {t.GetType()}");

            t = true;
            Console.WriteLine($"t.GetType(): {t.GetType()}");

            t = new List<int>();
            Console.WriteLine($"t.GetType(): {t.GetType()}");
        }

        static void InvokeMemberOnDynamicData()
        {
            dynamic textData1 = "Hello";

            try
            {
                Console.WriteLine(textData1.ToUpper());
                Console.WriteLine(textData1.toupper());
                Console.WriteLine(textData1.Foo(10, "type", true));
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ex)
            {
                Console.Write("Error => ");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
