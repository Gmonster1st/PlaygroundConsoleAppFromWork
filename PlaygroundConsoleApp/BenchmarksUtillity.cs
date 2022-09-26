using BenchmarkDotNet.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlaygroundConsoleApp
{
    [MemoryDiagnoser]
    [Orderer(BenchmarkDotNet.Order.SummaryOrderPolicy.FastestToSlowest)]
    [RankColumn]
    public class BenchmarksUtillity
    {
        private const string number = "1509";
        private const string word = "ANNA";
        private static readonly IntegerParser integerParser = new IntegerParser();
        private static readonly InterviewClass interviewClass = new InterviewClass();

        [Benchmark]
        public void GetIntFromStringManually()
        {
            integerParser.ManuallParse(number);
        }

        [Benchmark]
        public void GetIntFromStringLibrary()
        {
            integerParser.LibraryParse(number);
        }
        
        [Benchmark]
        public void GetIntFromStringLibraryTryParse()
        {
            integerParser.LibraryTryParse(number);
        }
        
        [Benchmark]
        public void IsPalindromLoop()
        {
            interviewClass.isPalindromLoop(word);
        }
        
        [Benchmark]
        public void IsPalindromReverse()
        {
            interviewClass.isPalindromReverse(word);
        }

        [Benchmark]
        public void IsPalindromRecursicveDefaultIndex()
        {
            interviewClass.isPalindromRecursicveDefaultIndex(word);
        }

        [Benchmark]
        public void IsPalindromRecursicveSubString()
        {
            interviewClass.isPalindromRecursicveSubString(word);
        }
    }
}
