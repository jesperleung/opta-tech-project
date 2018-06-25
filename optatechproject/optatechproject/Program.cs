using System;
using System.Collections.Generic;

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
            // string to be passed around for SQL operations
            string conString = "";

            // loading cities into list for lookup
            HashSet<string> cities = new HashSet<string>();
            cities = FileIO.LoadCities();
            // loading provinces into list for lookup
            HashSet<string> provinces = new HashSet<string>();
            provinces = FileIO.LoadProvinces();
            // loading street suffixes into list for lookup
            HashSet<string> suffixes = new HashSet<string>();
            suffixes = FileIO.LoadSuffixes();

            Console.Write("Please type the filename of the input data file: ");
            string inputfilename = Console.ReadLine();
            // Console.WriteLine(inputfilename);


            // connect to database
            conString = DBIO.ConnectToDB();

            // load XLS file
            FileIO.LoadXLS(inputfilename, cities, provinces, suffixes, conString);

            
            Console.WriteLine("All done. Press any key to finish...");
            Console.ReadKey(true);



        }
    }
}
