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
            // loading cities into list for lookup
            List<string> cities = new List<string>();
            cities = FileIO.LoadCities();
            // loading provinces into list for lookup
            List<string> provinces = new List<string>();
            provinces = FileIO.LoadProvinces();
            // loading street suffixes into list for lookup
            List<string> suffixes = new List<string>();
            suffixes = FileIO.LoadSuffixes();

            Console.Write("Please type the filename of the input data file: ");
            string inputfilename = Console.ReadLine();
            // Console.WriteLine(inputfilename);

            FileIO.LoadXLS(inputfilename, cities, provinces, suffixes);

            // connect to database
            //ConnectToDB();

            Console.WriteLine("All done. Press any key to finish...");
            Console.ReadKey(true);



        }
    }
}
