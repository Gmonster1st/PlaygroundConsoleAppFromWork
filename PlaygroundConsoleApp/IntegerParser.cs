using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlaygroundConsoleApp
{
    public class IntegerParser
    {
        public int ManuallParse(string s)
        {
            int sh = 0;
            for (int x = 0; x < s.Length; x++)
            {
                sh = sh * 10 + (s[x] - '0');
            }
            return sh;
        }

        public int LibraryParse(string s)
        {
            return int.Parse(s);
        }

        public int LibraryTryParse(string s)
        {
            return int.TryParse(s, out int result)? result : throw new Exception();
        }
    }
}
