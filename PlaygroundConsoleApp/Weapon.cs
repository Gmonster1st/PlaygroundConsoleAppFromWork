namespace PlaygroundConsoleApp
{
    public class Weapon
    {
        public string Name { get; set; }

        public void Label()
        {
            Console.WriteLine("This is {0}", Name);
        }
    }
}
