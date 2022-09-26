using static PlaygroundConsoleApp.Exercise;

namespace PlaygroundConsoleApp
{
    public class Run
    {
        public static void RunExercise()
        {
            Phone one = new Phone();
            Phone two = new Phone("Apple", "IPhone 12");
            Phone three = new Phone("Apple", "IPhone 12", "September 24, 2021");

            one.Introduce();
            two.Introduce();
            three.Introduce();
        }

    }

}
