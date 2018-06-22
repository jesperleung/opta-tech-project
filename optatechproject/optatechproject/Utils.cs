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
    }
}
