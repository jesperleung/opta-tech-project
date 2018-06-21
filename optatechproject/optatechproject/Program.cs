using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.IO;
// using Excel = Microsoft.Office.Interop.Excel;
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
        private static List<string> cities;
        private static List<string> provinces;

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

        //public static void LoadRawData(string filename)
        //{
        //    Excel.Application xlApp = new Excel.Application();
        //    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"./" + filename);
        //    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        //    Excel.Range xlRange = xlWorksheet.UsedRange;

        //    int rowCount = xlRange.Rows.Count;
        //    int colCount = xlRange.Columns.Count;

        //    for (int i = 1; i <= rowCount; i++)
        //    {
        //        for (int j = 1; j <= colCount; j++)
        //        {
        //            //new line
        //            if (j == 1)
        //                Console.Write("\r\n");

        //            //write the value to the console
        //            if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
        //                Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

        //            //add useful things here!   
        //        }
        //    }

        //    //cleanup
        //    GC.Collect();
        //    GC.WaitForPendingFinalizers();

        //    //rule of thumb for releasing com objects:
        //    //  never use two dots, all COM objects must be referenced and released individually
        //    //  ex: [somthing].[something].[something] is bad

        //    //release com objects to fully kill excel process from running in the background
        //    Marshal.ReleaseComObject(xlRange);
        //    Marshal.ReleaseComObject(xlWorksheet);

        //    //close and release
        //    xlWorkbook.Close();
        //    Marshal.ReleaseComObject(xlWorkbook);

        //    //quit and release
        //    xlApp.Quit();
        //    Marshal.ReleaseComObject(xlApp);

        //}

        // load and read CSV file with provided filename ***FILE MUST BE IN ROOT FOLDER OF PROJECT****
        public static void LoadCSV(string filename)
        {
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
                            // replace any non-alphanumeric and non-space characters with a space
                            string clean = Regex.Replace(field, "[^A-Za-z0-9 ]", " ");
                            ParseAddress(clean);
                            //Console.WriteLine(clean);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                // if error, print error to console
                Console.WriteLine(e.ToString());
            }
        }
        // loads list of Canadian cities from csv file into a List and returns it
        public static List<string> LoadCities()
        {
            // initialize List for cities
            List<string> cities = new List<string>();

            try
            {
                // getting data file from root folder of project and provided filename
                using (TextFieldParser parser = new TextFieldParser(@"..\..\..\..\places.csv"))
                {
                    Console.WriteLine("loading cities data...");
                    // telling parser to read a CSV
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");
                    while (!parser.EndOfData)
                    {
                        // read through every line in CSV file
                        string[] fields = parser.ReadFields();
                        foreach (string field in fields)
                        {
                            // add city name to List
                            cities.Add(field);
                        }
                    }
                }
                Console.WriteLine("** done loading cities **");
            }
            catch (Exception e)
            {
                // if error, print error to console
                Console.WriteLine(e.ToString());
            }

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
        // parses raw string into its parts
        public static void ParseAddress(string raw)
        {
            // initialize counter to be used to ensure all elements are gathered from raw string
            int counter;

            // split raw string into parts about the spaces
            string[] split = Regex.Split(raw, @"\s+");
            // convert array of substrings into List
            List<string> converted = new List<string>(split);

            // for each element in split string
            foreach (string s in split)
            {
                // reset counter to 0 for every iteration
                counter = 0;
                // check if current substring exists in list of provinces
                if (provinces.Contains(s))
                {
                    Console.WriteLine("provinces exists");
                    counter++;
                }
                // check if current substring exists in list of cities
                else if (cities.Contains(s))
                {
                    Console.WriteLine("city exists");
                    counter++;
                }
                // check if current substring is of the postal code format ex. A1A1A1
                else if (Regex.IsMatch(s, @"\w\d\w\d\w\d"))
                {
                    Console.WriteLine("postal code exists");
                    counter++;
                }
                // check if current substring is all numeric (street numbers are all numeric)
                else if (IsNumeric(s))
                {
                    Console.WriteLine("street number exists");
                    counter++;
                }
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

            Console.Write("Please type the filename of the input data file: ");
            string inputfilename = Console.ReadLine();
            Console.WriteLine(inputfilename);

            LoadCSV(inputfilename);

            // connect to database
            //ConnectToDB();

            Console.WriteLine("All done. Press any key to finish...");
            Console.ReadKey(true);



        }
    }
}
