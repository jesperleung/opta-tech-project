using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OptaTechProject
{
    class Utils
    {
        // iterates through a string to check if each character is a number
        public static bool IsNumeric(string str)
        {
            // get each char c in string str
            foreach (char c in str)
            {
                // if char c is not a number, return false
                if (c < '0' || c > '9')
                    return false;
            }
            // else return true
            return true;
        }
        // gets all substrings and stores it in a list ex. ABC -> A, AB, ABC, BC, C
        // code adapted from https://www.dotnetperls.com/all-substrings
        public static List<string> AllSubstrings(string[] original)
        {
            List<string> substrings = new List<string>();
            

            // Avoid full length.
            for (int length = 1; length <= original.Length; length++)
            {
                // End index is tricky.
                for (int start = 0; start <= original.Length - length; start++)
                {
                    string substring = String.Join(" ", original, start, length);
                    substrings.Add(substring);
                }
            }
            return substrings;
        }
    }
}
