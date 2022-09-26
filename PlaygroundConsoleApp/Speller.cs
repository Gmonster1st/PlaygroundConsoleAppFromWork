namespace PlaygroundConsoleApp
{
    public class Speller
    {
        public static string Convert(int i)
        {
            Dictionary<int, string> numbers = new Dictionary<int, string>()
            {
                { 0,"zero"},
                { 1,"one"},
                { 2,"two"},
                { 3,"three"},
                { 4,"four"},
                { 5,"five"}
            };

            if (numbers.ContainsKey(i))
            {
                return numbers[i];
            }
            else
            {
                return "nope";
            }
        }
    }
}


