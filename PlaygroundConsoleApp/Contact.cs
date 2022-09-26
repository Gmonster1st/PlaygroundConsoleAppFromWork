namespace PlaygroundConsoleApp
{
    public class Contact
    {
        public string Name { get; set; }
        public string PhoneNumber { get; set; }

        public Contact(string name, string phoneNumber)
        {
            Name = name;
            PhoneNumber = phoneNumber;
        }

        public void Call()
        {
            System.Console.WriteLine("Calling to {0}, Phone number is {1}", Name, PhoneNumber);
        }
    }
}
