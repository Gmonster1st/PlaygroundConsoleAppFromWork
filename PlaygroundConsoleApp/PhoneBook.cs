using System.Collections;

namespace PlaygroundConsoleApp
{
    public class PhoneBook : IEnumerable<Contact>
    {
        public List<Contact> Contacts { get; private set; }

        public PhoneBook()
        {
            Contacts = new List<Contact>{
                new Contact("Andre", "435797087"),
                new Contact("Andre", "435797087"),
                new Contact("Andre", "435797087"),
                new Contact("Andre", "435797087")
            };
        }
        public IEnumerator<Contact> GetEnumerator()
        {
            return Contacts.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return Contacts.GetEnumerator();
        }
    }
}
