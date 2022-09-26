namespace PlaygroundConsoleApp
{
    public class Gun : Weapon, IShootable
    {
        public Gun()
        {
            base.Name = "Gun";
        }

        public void Shoot()
        {
            Console.WriteLine("Bang");
        }
    }
}
