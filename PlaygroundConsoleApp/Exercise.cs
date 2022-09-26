namespace PlaygroundConsoleApp
{
    public partial class Exercise
    {
        public static void Check(int number)
        {
            string result = ((number % 2) == 0) ? "Even" : "Odd";
            Console.WriteLine(result);
        }
        public static void NestedCheck(int number)
        {
            if ((number % 2) != 0)
            {
                if ((number % 3) == 0)
                {
                    Console.WriteLine("Divisible by {0}", 3);
                }
                else
                {
                    Console.WriteLine("Odd number");
                }
                
            }
            else
            {
                if ((number % 7) == 0)
                {
                    Console.WriteLine("Divisible by {0}", 7);
                }
                else
                {
                    Console.WriteLine("Even number");
                }
            }
        }

        public static void ForLoop()
        {
            for (int i = -3; i <= 3; i++)
            {
                Console.WriteLine(i);
            }
        }

        public static void WhileLoop()
        {
            int i = 3;
            while (i >= -3)
            {
                Console.WriteLine(i);
            }
        }

        public static void infiniteLoop()
        {
            int i = -10;

            while (true)
            {

                if (i % 3 == 0)
                {
                    i++;
                    continue;
                }

                if (i == 10) break;

                Console.WriteLine(i++);
            }
        }

        public static void GetOdd(int[] Array)
        {
            foreach (var item in Array)
            {
                if ((item % 2) != 0) Console.WriteLine(item);
            }
        }

        public static void GetEven(int[] Array)
        {
            foreach (var item in Array)
            {
                if ((item % 2) == 0) Console.WriteLine(item);
            }
        }

        public static void Run()
        {
            WhileLoop();
            ForLoop();

            int[] array = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };

            GetOdd(array);
            GetEven(array);

        }

    }
}
