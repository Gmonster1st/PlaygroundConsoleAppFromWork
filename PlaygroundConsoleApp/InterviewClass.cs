using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlaygroundConsoleApp
{
    public class InterviewClass
    {
        public bool isPalindromLoop(string word)
        {
            int j = word.Length - 1;
            for (int i = 0; i < word.Length; i++)
            {
                if (word[i] != word[j]) return false;
                
                if (i >= j) break;

                j--;                                
            }

            return true;
        }
        
        public bool isPalindromReverse(string word)
        {
            //alter
            var rev = word.Reverse();
            var res = rev.Equals(word);
            return res;
        }

        public bool isPalindromRecursicveDefaultIndex(string word, int index = 0)
        {
            //Get the word length
            //Check if the leght is more that 0
            //Compare if the 1st and last letters are Equal and if the index is in the midle return result
            // otherwise move to the next call of self with index +1
            //
            int length = word.Length;
            if (word.Length > 0)
            {
                return (word.Length == 1) ? true : (index >= (length - 1 - index)) ? true : (word[index] == word[length - 1 - index])? isPalindromRecursicveDefaultIndex(word, ++index): false;
            }
            return false;
        }
        
        public bool isPalindromRecursicveSubString(string word)
        {
            //Check if 1st and last letters are equal and continue with new subString else return
            return (word.Length > 2) && (word[0] == word[word.Length - 1]) ? isPalindromRecursicveSubString(word.Substring(1, word.Length-2)) : (word[0] == word[word.Length - 1]);
        }

        public string stringMethodB(string blabla)
        {
            blabla = "cancelled";
            return blabla;
        }

        public string stringMethodA(string blabla)
        {
            var result = stringMethodB(blabla);
            return blabla;
        }

        public void run()
        {
            var testResult = stringMethodA("ciao");

            Console.WriteLine(testResult);
        }
    }
}
