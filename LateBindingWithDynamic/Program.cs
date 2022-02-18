using System;
using System.Reflection;

namespace LateBindingWithDynamic
{
    namespace LateBindingWithDynamic
    {
        class Program
        {
            static void Main(string[] args)
            {
                Assembly asm = Assembly.LoadFrom("c:\\Users\\Greed\\Desktop\\C# .NET\\DLR\\MathLibrary\\bin\\Debug\\net5.0\\MathLibrary.dll");

                Console.WriteLine("****** Late Binding With Dynamic ******");
                Console.WriteLine();

                AddWithReflection(asm);
                Console.WriteLine();

                AddWithDynamic(asm);
                Console.WriteLine();
            }

            static void AddWithReflection(Assembly asm)
            {
                Console.WriteLine("Used Reflection => ");


                try
                {
                    Type upClass = asm.GetTypes()[0];

                    object obj = asm.CreateInstance(upClass.FullName);

                    MethodInfo method = upClass.GetMethod("Add");

                    object[] args = { 10, 10 };

                    Console.WriteLine($"{upClass.FullName}.{method.Name}(10, 10): {method.Invoke(obj, args)}");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                
            }

            static void AddWithDynamic(Assembly asm)
            {
                Console.WriteLine("Used Dynamic => ");

                try
                {
                    string className= asm.GetTypes()[0].FullName;

                    dynamic obj = asm.CreateInstance(className);

                    Console.WriteLine($"{className}.Add(10, 10): {obj.Add(10, 10)}");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
    }
}
