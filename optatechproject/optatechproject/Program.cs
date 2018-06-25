using System;
using System.Collections.Generic;
using System.Data.SqlClient;

/*
jesper leung
june 19th 2018
opta information intelligence technical project
*/

namespace OptaTechProject
{
    public class OptaTechProject
    {
        public static void Main(string[] args)
        {
            // conString to be passed around
            string conString;
            // loading cities into list for lookup
            HashSet<string> cities = new HashSet<string>();
            cities = FileIO.LoadCities();
            // loading provinces into list for lookup
            HashSet<string> provinces = new HashSet<string>();
            provinces = FileIO.LoadProvinces();
            // loading street suffixes into list for lookup
            HashSet<string> suffixes = new HashSet<string>();
            suffixes = FileIO.LoadSuffixes();

            conString = DBIO.ConnectToDB();

            
            Console.Write("Please type the filename of the input data file: ");
            string inputfilename = Console.ReadLine();
            // Console.WriteLine(inputfilename);

            // print headings for tabulated display
            Console.WriteLine("{0, -15} {1, -25} {2, -40} {3, -10} {4, -15}", "Street #", "Street Name", "City", "Province", "Postal Code");

            FileIO.LoadXLS(inputfilename, cities, provinces, suffixes, conString);

   
            Console.WriteLine("All done. Press any key to finish...");
            Console.ReadKey(true);



        }
    }
}
