using System;
 
namespace ProgrammingHub
{
    class BinaryTriangle
    {
        public static void Main()
        {
            int p, lastInt = 0, input;
 
            Console.Write("Enter the Number of Rows : ");
            input = int.Parse(Console.ReadLine());
 
            Console.WriteLine("\nBinary Triangle  \n");
 
            for (int i = 1; i <= input; i++)
            {
                for (p = 1; p <= i; p++)
                {
                    if (lastInt == 1)
                    {
                        Console.Write("0");
                        lastInt = 0;
                    }
                    else if (lastInt == 0)
                    {
                        Console.Write("1");
                        lastInt = 1;
                    }
                }
                Console.Write("\n");
            }
            Console.ReadLine();
        }
    }
}