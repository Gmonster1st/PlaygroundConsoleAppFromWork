namespace PlaygroundConsoleApp
{
    public partial class Exercise
    {
        public class Phone
        {
            public string Company;
            public string Model;
            public string ReleaseDay;

            // Place for your constructors
            public Phone(string company = "Unknown", string model = "Unknown", string releaseDay = "Unknown")
            {
                Company = company;
                Model = model;
                ReleaseDay = releaseDay;
            }


            public void Introduce()
            {
                Console.WriteLine("It is {0} created by {1}. It was released {2}.", Model, Company, ReleaseDay);
            }

        }

    }
}
