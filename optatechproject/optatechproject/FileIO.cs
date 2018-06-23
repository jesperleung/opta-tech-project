using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace OptaTechProject
{
    class FileIO
    {
        /*
        * load and read XSL file with provided filename ***FILE MUST BE IN ROOT FOLDER OF PROJECT***
        * code adapted from https://coderwall.com/p/app3ya/read-excel-file-in-c
        */
        public static void LoadXSL(string filename, List<string> cities, List<string> provinces, List<string> suffixes)
        {
            try
            {
                // Creating Excel objects
                Excel.Application xlApp = new Excel.Application();
                filename = Path.GetFullPath("@..\\..\\..\\..\\..\\..\\" + filename);
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filename);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                //Console.WriteLine("row count: " + rowCount);
                int colCount = xlRange.Columns.Count;
                //Console.WriteLine("column count: " + colCount);

                string raw;
                Console.WriteLine("{0, -15} {1, -25} {2, -40} {3, -10} {4, -15}", "Street #", "Street Name", "City", "Province", "Postal Code");
                //iterate over the rows and columns and print to the console as it appears in the file
                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value != null)
                        {
                            raw = xlRange.Cells[i, j].Value;
                            ParseAddress(raw, cities, provinces, suffixes);
                        }

                    }
                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        // load and read CSV file with provided filename ***FILE MUST BE IN ROOT FOLDER OF PROJECT****
        public static List<string> LoadCSV(string filename)
        {
            List<string> loaded = new List<string>();
            try
            {
                // getting data file from root folder of project and provided filename
                using (TextFieldParser parser = new TextFieldParser(@"..\..\..\..\" + filename))
                {
                    // telling parser to read a CSV
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");
                    while (!parser.EndOfData)
                    {
                        // read through every line in CSV file
                        string[] fields = parser.ReadFields();
                        foreach (string field in fields)
                        {
                            loaded.Add(field);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                // if error, print error to console
                Console.WriteLine(e.ToString());
            }
            return loaded;
        }
        // loads list of Canadian cities from csv file into a List and returns it
        public static List<string> LoadCities()
        {
            // initialize List for cities
            List<string> cities = new List<string>();

            cities = LoadCSV("places.csv");

            return cities;
        }
        // loads list of Canadian provinces into a List and returns it
        public static List<string> LoadProvinces()
        {
            List<string> provinces = new List<string>
            {
                "ON",
                "QC",
                "BC",
                "AB",
                "MB",
                "NL",
                "PE",
                "NS",
                "NB",
                "SK",
                "YT",
                "NT",
                "NU"
            };

            return provinces;
        }
        // loads list of street suffixes into a List and returns it
        public static List<string> LoadSuffixes()
        {
            List<string> suffixes = new List<string>();

            suffixes = LoadCSV("street suffixes.csv");

            return suffixes;
        }
        // parses raw string into its parts
        public static void ParseAddress(string raw, List<string> cities, List<string> provinces, List<string> suffixes)
        {
            // counter to ensure all components have been found
            int counter = 0;
            // variables to store address components
            string streetnum = "";
            string streetname = "";
            string city = "";
            string province = "";
            string postalcode = "";

            // replace anything that isn't alphanumeric but keeping periods
            raw = Regex.Replace(raw, @"[!@#$%^&*()_+=\[{\]};:<>|/?,\\""]", " ");

            // split raw string into parts about the spaces
            string[] split = Regex.Split(raw, @"\s+");

            // convert array of substrings into List
            List<string> converted = new List<string>(split);

            // removing empty spaces in converted
            foreach (var x in converted.ToList())
            {
                if (String.IsNullOrWhiteSpace(x))
                {
                    converted.Remove(x);
                }
            }

            // street number is first element in the list
            streetnum = converted.ElementAt(0);
            converted.Remove(streetnum);
            counter++;
            // postal code is last element in the list
            postalcode = converted.ElementAt(converted.Count - 1);
            converted.Remove(postalcode);
            counter++;
            // province is 2nd last element in the list
            province = converted.ElementAt(converted.Count - 1);
            converted.Remove(province);
            counter++;

            foreach (var x in converted)
            {
                if (suffixes.Contains(Regex.Replace(String.Join(" ", converted.ToArray(), converted.IndexOf(x), 1), @"[.]", "")))
                {
                    // SPECIAL CASE: multiple occurences of street suffixes i.e. BEACH RD, or xxx ST. ST. John's
                    // check the next element after x to see if it is also a street - if not, proceed as usual, otherwise skip to the next element 
                    if (!suffixes.Contains(Regex.Replace(String.Join(" ", converted.ToArray(), converted.IndexOf(x) + 1, 1), @"[.]", "")))
                    {
                        // street suffix found, street name is everything before
                        streetname = String.Join(" ", converted.ToArray(), 0, converted.IndexOf(x) + 1);
                        counter++;
                        // and city is everything after
                        city = String.Join(" ", converted.ToArray(), converted.IndexOf(x) + 1, converted.Count - converted.IndexOf(x) - 1);
                        counter++;
                        // stop iterating in case there's another street suffix in the city name
                        break;
                    }
                }
            }

            // all information extracted, print to screen
            Console.WriteLine("{0, -15} {1, -25} {2, -40} {3, -10} {4, -15}", streetnum, streetname, city, province, postalcode);


        }
    }
}
