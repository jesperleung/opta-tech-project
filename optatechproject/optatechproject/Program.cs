using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic.FileIO;
using System.Text.RegularExpressions;

/*
jesper leung
june 19th 2018
opta information intelligence technical project
*/

namespace OptaTechProject
{
    public class OptaTechProject
    {
        //creates SQL connection string and tries to connect to the SQL Server on localhost called SQLEXPRESS
        public static void ConnectToDB()
        {
            try
            {
                // connection string from Server Explorer in VS after adding database
                string conString = "Data Source=.\\SQLEXPRESS;Initial Catalog=optatechproject;Integrated Security=True";
                SqlConnection con = new SqlConnection(conString);

                Console.Write("Connecting to SQL Server ... ");

                con.Open();
                Console.WriteLine("Done.");
            }
            catch (SqlException e)
            {
                // if error, print error to console
                Console.WriteLine(e.ToString());
            }
        }

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

                //iterate over the rows and columns and print to the console as it appears in the file
                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value != null)
                        {
                            raw = xlRange.Cells[i, j].Value;
                            ParseAddress(raw, cities, provinces);
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
        public static void ParseAddress(string raw, List<string> cities, List<string> provinces)
        {
            // initialize counter to be used to ensure all elements are gathered from raw string
            int counter;
            // variables to store address components
            string streetnum;
            string streetname;
            string city;
            string province;
            string postalcode;
            // used for checking for all substrings in street + city combination
            string substring;
            Boolean cityfound = false;
            // replace anything that isn't alphanumeric or a symbol that could be used with " "
            raw = Regex.Replace(raw, @"[!@#$%^&*()_+=\[{\]};:<>|/?,\\""]", " ");
            // split raw string into parts about the spaces
            string[] split = Regex.Split(raw, @"\s+");
            // convert array of substrings into List
            List<string> converted = new List<string>(split);

            // reset counter to 0 for every iteration
            counter = 0;
            //Console.WriteLine("Before: "+converted.Count);
            // for each element in split string
            foreach (var s in converted.ToList())
            {
                // check if current substring is of the postal code format ex. A1A1A1
                //Console.WriteLine(s);
                if (Regex.IsMatch(s, @"\w\d\w\s*\d\w\d"))
                {
                    // save postal code
                    postalcode = s;
                    converted.Remove(s);
                    //Console.WriteLine("postal code exists: " + s);
                    counter++;
                }
                // check if current substring is all numeric (street numbers are all numeric)
                else if (IsNumeric(s))
                {
                    // save street number
                    streetnum = s;
                    converted.Remove(s);
                    //Console.WriteLine("street number exists: " + s);
                    counter++;
                }
                // check if current substring exists in List of provinces
                else if (provinces.Contains(s, StringComparer.OrdinalIgnoreCase))
                {
                    // save province
                    province = s;
                    converted.Remove(s);
                    //Console.WriteLine("provinces exists:" + s);
                    counter++;
                }
                // remove any whitespace, null, empty entries from list
                else if (string.IsNullOrWhiteSpace(s))
                {
                    converted.Remove(s);
                }
                //Console.WriteLine("After: "+converted.Count);
            }
            // if all have been accounted for so far (street #, province, postal code)
            if (counter >= 3)
            {
                // combine street name and city name into one string again
                split = converted.ToArray();
                Console.WriteLine("Street name and city: " + String.Join(" ", split));
                // get all substrings of remaining text to find city name IN ORDER ex. A, AB, ABC, BC, C
                for (int length = 1; length < split.Length; length++)
                { 
                    for (int start = 0; start <= split.Length - length; start++)
                    {
                        // get substrings of words in remaining text
                        substring = String.Join(" ", split, start, length);

                        //Console.WriteLine(substring);
                        // if substring is name of a city
                        if (cities.Contains(substring, StringComparer.OrdinalIgnoreCase))
                        {
                            // save city name
                            Console.WriteLine("city found: " + substring);
                            city = substring;
                            cityfound = true;

                            // city is at the start of the string
                            if (start == 0)
                            {
                                // get everything after the city name as streetname
                                streetname = String.Join(" ", split, length, split.Length - length);
                            }
                            // city ends at the end of the string
                            else
                            {
                                // get everything before the city name as streetname
                                streetname = String.Join(" ", split, 0, start);
                            }
                            Console.WriteLine("street name: " + streetname);
                            // stop the for loop once city has been found
                            break;
                        }
                    }
                }
                // city couldn't be found in List, need user input ***SHOULDN'T HAPPEN OFTEN***
                if (!cityfound)
                {
                    Console.WriteLine("Address could not be parsed properly, user input required");
                    Console.WriteLine(String.Join(" ", split));
                    Console.WriteLine("What is the street name?");
                    streetname = Console.ReadLine();
                    Console.WriteLine("What is the city?");
                    city = Console.ReadLine();
                }
            }
            // not enough elements matched, need user input ***SHOULDN'T HAPPEN OFTEN***
            else
            {
                Console.WriteLine("Address could not be parsed properly, user input required");
                Console.WriteLine(raw);
                Console.WriteLine("What is the street number?");
                streetnum = Console.ReadLine();
                Console.WriteLine("What is the street name?");
                streetname = Console.ReadLine();
                Console.WriteLine("What is the city?");
                city = Console.ReadLine();
                Console.WriteLine("What is the province?");
                province = Console.ReadLine();
                Console.WriteLine("What is the postal code?");
                postalcode = Console.ReadLine();

            }




        }
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

        public static void Main(string[] args)
        {
            // loading cities into list for lookup
            List<string> cities = new List<string>();
            cities = LoadCities();
            // loading provinces into list for lookup
            List<string> provinces = new List<string>();
            provinces = LoadProvinces();
            // loading street suffixes into list for lookup
            List<string> suffixes = new List<string>();
            suffixes = LoadSuffixes();

            Console.Write("Please type the filename of the input data file: ");
            string inputfilename = Console.ReadLine();
            // Console.WriteLine(inputfilename);

            LoadXSL(inputfilename, cities, provinces, suffixes);

            // connect to database
            //ConnectToDB();

            Console.WriteLine("All done. Press any key to finish...");
            Console.ReadKey(true);



        }
    }
}
