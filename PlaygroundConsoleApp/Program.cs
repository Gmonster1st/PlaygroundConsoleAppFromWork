using BenchmarkDotNet.Running;
using System.Configuration;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace PlaygroundConsoleApp
{
    public class Program
    {
        private static readonly log4net.ILog log = LogHelper.GetLogger();

        private static void Main(string[] args)
        {
            string stringForFloat = "0.85"; // datatype should be float
            string stringForInt = "12345"; // datatype should be int

            float parsedFloat;
            string floatStatus = float.TryParse(stringForFloat, out parsedFloat) ? "Parsing of float succedeed" : "Parsing of float failed";

            int parsedInt;
            string intStatus = int.TryParse(stringForInt, out parsedInt) ? "Parsing of int succedeed" : "Parsing of int failed";

            Console.WriteLine(floatStatus);
            Console.WriteLine(intStatus);
            Method.Run();

            Exercise.NestedCheck(3);
            Exercise.NestedCheck(28);
            Exercise.NestedCheck(2);
            Exercise.NestedCheck(5);

            Exercise.infiniteLoop();

            var test = new InterviewClass();

            test.run();

            // new instance
            Gun pist = new Gun();

            // test for methods
            pist.Label();
            pist.Shoot();

            // verifying the interface and the parent class
            if (pist is IShootable && pist is Weapon)
                Console.WriteLine("Yes, it is my parents.");

            //BenchmarkRunner.Run<BenchmarksUtillity>();
            log4net.Config.XmlConfigurator.Configure();
            log.Error("This is my error message!");
        }

        //Class myClass = new Class();
        //Console.WriteLine(myClass.isPalindromRecursicveDefaultIndex("ANNA"));
        //Console.WriteLine(myClass.isPalindromRecursicveDefaultIndex("ANIA"));
        //Console.WriteLine(myClass.isPalindromRecursicveSubString("ANIA"));
        //Console.WriteLine(myClass.isPalindromRecursicveSubString("ANNA"));
    }
}