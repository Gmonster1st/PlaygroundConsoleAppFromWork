using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlaygroundConsoleApp
{
    public class Method
    {
        // Place for your methods

        private static void Count(string s)
        {
            Console.WriteLine(s.Length);
        }

        private static string LowUpper(string s)
        {
            if (!string.IsNullOrEmpty(s))
            {
                return s.ToLower() + s.ToUpper();
            }

            throw new ArgumentNullException("LowUpper", "The argument of the method can not be null");
        }


        public static void Run()
        {
            // We encourage you to test your code with different strings,
            // but don't forget to put the default string back at the end of your testing.
            string s = "HeY ThErE !";

            /// Change nothing down here 
            s = LowUpper(s);
            Console.WriteLine(s);
            Count(s);

        }

    }
}
