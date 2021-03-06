﻿using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
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
        public static void LoadXLS(string filename, HashSet<string> cities, HashSet<string> provinces, HashSet<string> suffixes, string conString)
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
                            DBIO.WriteRaw(conString, raw);
                            ParseAddress(raw, cities, provinces, suffixes, conString);
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
        public static HashSet<string> LoadCSV(string filename)
        {
            HashSet<string> loaded = new HashSet<string>();
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
                            loaded.Add(Utils.RemoveDiacritics(field));
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
        public static HashSet<string> LoadCities()
        {
            // initialize List for cities
            HashSet<string> cities = new HashSet<string>();

            cities = LoadCSV("cities.csv");

            return cities;
        }
        // loads list of Canadian provinces into a List and returns it
        public static HashSet<string> LoadProvinces()
        {
            HashSet<string> provinces = new HashSet<string>
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
        public static HashSet<string> LoadSuffixes()
        {
            HashSet<string> suffixes = new HashSet<string>();

            suffixes = LoadCSV("street suffixes.csv");

            return suffixes;
        }
        // parses raw string into its parts
        public static void ParseAddress(string raw, HashSet<string> cities, HashSet<string> provinces, HashSet<string> suffixes, string conString)
        {
            // initialize counter to be used to ensure all elements are gathered from raw string
            int counter;

            // variables to store address components
            string streetnum;
            string streetname;
            string city;
            string province = "";
            string postalcode = "";

            string potentialstreet;
            string potentialcity;

            // used to end loops early
            Boolean cityfound = false;
            Boolean numfound = false;

            // pointers to use for finding street and city names
            int suffixindex = 0;
            int numindex = 0;
            int cityindex = 0;
            int pointer = 0;
            int j;

            // replace anything that isn't alphanumeric or a symbol that could be used with " "
            raw = Regex.Replace(raw, @"[!@#$%^&*()_+=\[{\]};:<>|/?,\\""]", " ");

            // split raw string into parts about the spaces
            string[] split = Regex.Split(raw, @"\s+");

            // convert array of substrings into List
            List<string> converted = new List<string>(split);

            // used to store results from AllSubstrings method
            List<string> substrings = new List<string>();

            // reset counter to 0 for every iteration
            counter = 0;
            //Console.WriteLine("Before: "+converted.Count);
            // for each element in split string
            foreach (var s in converted.ToList())
            {
                // check if current substring is of the postal code format ex. A1A1A1
                if (Regex.IsMatch(s, @"\w\d\w\s*\d\w\d"))
                {
                    // Console.WriteLine(s);
                    // save postal code
                    postalcode = s;
                    converted.Remove(s);
                    //Console.WriteLine("postal code exists: " + s);
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
            // if province and postal code have been accounted for so far
            if (counter >= 2)
            {
                //Console.WriteLine("Street # Street name and city: " + String.Join(" ", split));
                // for each element in converted (made up of street number, street name, city name)
                foreach (var x in converted)
                {
                    // first find street suffix
                    if (suffixes.Contains(x, StringComparer.OrdinalIgnoreCase))
                    {
                        // get index of street suffix
                        suffixindex = converted.IndexOf(x);
                        // CASE 1: suffix is at the end of the string, so street name is at least 1 element before it
                        if (suffixindex == (converted.Count - 1))
                        {
                            potentialstreet = converted.ElementAt(suffixindex);
                            potentialcity = "";
                            pointer = suffixindex - 1;
                            // start building street name from right to left until search reaches a #
                            while (!numfound && !cityfound && pointer != 0)
                            {
                                // building string from right to left
                                potentialstreet = converted.ElementAt(pointer) + " " + potentialstreet;
                                pointer--;
                                //Console.WriteLine("building string backwards: " + potentialstreet);
                                // if next element to the left is a number, could be street number
                                if (Utils.IsNumeric(converted.ElementAt(pointer)))
                                {
                                    numindex = pointer;
                                    // only 1 element before numindex, check that against cities
                                    if (numindex == 1)
                                    {
                                        potentialcity = String.Join(" ", converted.GetRange(0, numindex).ToArray());
                                        // first element in string is city, so everything from numindex to suffixindex is the street number and name
                                        if (cities.Contains(Utils.RemoveDiacritics(potentialcity), StringComparer.OrdinalIgnoreCase))
                                        {
                                            city = potentialcity;
                                            streetnum = converted.ElementAt(numindex);
                                            streetname = String.Join(" ", converted.GetRange(numindex + 1, suffixindex - numindex));
                                            numfound = true;
                                            cityfound = true;
                                            Console.WriteLine("{0, -15} {1, -25} {2, -40} {3, -10} {4, -15}", streetnum, streetname, city, province, postalcode);
                                        }
                                    }
                                    // street number is first element in string, so city must be between 1 and at most suffixindex - 1
                                    else if (numindex == 0)
                                    {
                                        // get all substrings from 1 to current potentialstreet
                                        substrings = Utils.AllSubstrings(converted.GetRange(1, suffixindex - 1).ToArray());
                                        // check every substring for a city name
                                        foreach (var y in substrings)
                                        {
                                            // one of the substrings is a city name
                                            if (cities.Contains(Utils.RemoveDiacritics(y), StringComparer.OrdinalIgnoreCase))
                                            {
                                                potentialcity = y;
                                                //Console.WriteLine("potential city: " + potentialcity);
                                            }
                                        }
                                        // last iteration of potentialcity should be the city name (longest one)
                                        city = potentialcity;
                                        cityindex = city.Split(null).Length;
                                        streetnum = converted.ElementAt(numindex);
                                        streetname = String.Join(" ", converted.GetRange(cityindex + 1, suffixindex - cityindex));
                                        Console.WriteLine("{0, -15} {1, -25} {2, -40} {3, -10} {4, -15}", streetnum, streetname, city, province, postalcode);
                                        cityfound = true;
                                        numfound = true;
                                    }
                                    // multiple elements before numindex, need to find cities
                                    else
                                    {
                                        // get all substrings from 0 to index of potential street number
                                        substrings = Utils.AllSubstrings(converted.GetRange(0, numindex).ToArray());
                                        // check every substring for a city name
                                        foreach (var y in substrings)
                                        {
                                            // one of the substrings is a city name
                                            if (cities.Contains(Utils.RemoveDiacritics(y), StringComparer.OrdinalIgnoreCase))
                                            {
                                                potentialcity = y;
                                                // Console.WriteLine("potential city: " + potentialcity);
                                            }
                                        }
                                        // found a city name that spans from 0 to numindex - 1, so everything from numindex to suffixindex is the street number and name
                                        // .Split(null) assumes whitespace
                                        if (potentialcity.Split(null).Length == numindex)
                                        {
                                            city = potentialcity;
                                            cityfound = true;
                                            streetnum = converted.ElementAt(numindex);
                                            numfound = true;
                                            streetname = String.Join(" ", converted.GetRange(numindex + 1, suffixindex - numindex));
                                            Console.WriteLine("{0, -15} {1, -25} {2, -40} {3, -10} {4, -15}", streetnum, streetname, city, province, postalcode);
                                        }
                                    }
                                }
                            }
                        }
                        // CASE 2: suffix is in the middle of the string, so street name must be somewhere in 0 to suffixindex
                        else
                        {
                            potentialstreet = "";
                            potentialcity = "";
                            pointer = 0;

                            // start building street from left to right, starting from 0 to suffixindex or until a # is found
                            while (!numfound && !cityfound)
                            {
                                potentialstreet = potentialstreet + converted.ElementAt(pointer);
                                //Console.WriteLine("building string forwards: " + potentialstreet);

                                // if next element to the right is a number, could be street number
                                if (Utils.IsNumeric(converted.ElementAt(pointer)))
                                {
                                    numindex = pointer;
                                    // street number is first element in string, street name must be from 1 to suffixindex, and city from suffixindex + 1 to converted.Count
                                    if (numindex == 0)
                                    {
                                        // get all substrings from 1 to converted.Count
                                        substrings = Utils.AllSubstrings(converted.GetRange(suffixindex + 1, converted.Count - suffixindex - 1).ToArray());
                                        // check every substring for a city name
                                        foreach (var y in substrings)
                                        {
                                            // one of the substrings is a city name
                                            if (cities.Contains(Utils.RemoveDiacritics(y), StringComparer.OrdinalIgnoreCase))
                                            {
                                                potentialcity = y;
                                                //Console.WriteLine("potential city: " + potentialcity);
                                            }
                                        }
                                        // city not found
                                        if (potentialcity == "")
                                        {
                                            DBIO.WriteError(conString, raw);
                                            break;
                                        }
                                        else
                                        {
                                            // last iteration of potentialcity should be the city name (longest one)
                                            city = potentialcity;
                                            streetnum = converted.ElementAt(numindex);
                                            streetname = String.Join(" ", converted.GetRange(1, suffixindex));
                                            cityfound = true;
                                            numfound = true;
                                            Console.WriteLine("{0, -15} {1, -25} {2, -40} {3, -10} {4, -15}", streetnum, streetname, city, province, postalcode);
                                            DBIO.WriteComplete(conString, streetnum, streetname, city, province, postalcode);
                                        }

                                    }
                                    // street number is last element in string, street name and city are somewhere in 0 to suffixindex
                                    else if ((numindex == converted.Count - 1))
                                    {
                                        // street ending is 2nd last element in string, so street name and city name are somewhere between 0 and suffixindex - 1
                                        if (suffixindex == numindex - 1)
                                        {
                                            substrings = Utils.AllSubstrings(converted.GetRange(0, converted.Count - suffixindex - 1).ToArray());
                                            // check every substring for a city name
                                            foreach (var y in substrings)
                                            {
                                                // one of the substrings is a city name
                                                if (cities.Contains(Utils.RemoveDiacritics(y), StringComparer.OrdinalIgnoreCase))
                                                {
                                                    potentialcity = y;
                                                    //Console.WriteLine("potential city: " + potentialcity);
                                                }
                                            }
                                            // city not found
                                            if (potentialcity == "")
                                            {
                                                DBIO.WriteError(conString, raw);
                                                break;
                                            }
                                            else
                                            {
                                                // last iteration of potentialcity should be the city name (longest one)
                                                city = potentialcity;
                                                cityindex = city.Split(null).Length;
                                                streetnum = converted.ElementAt(numindex);
                                                streetname = String.Join(" ", converted.GetRange(cityindex, suffixindex - cityindex + 1));
                                                cityfound = true;
                                                numfound = true;
                                                Console.WriteLine("{0, -15} {1, -25} {2, -40} {3, -10} {4, -15}", streetnum, streetname, city, province, postalcode);
                                            }

                                        }
                                        // street ending is in the middle, so street name is 0 to suffixindex, and city is suffixindex + 1 onwards to numindex - 1
                                        else
                                        {
                                            substrings = Utils.AllSubstrings(converted.GetRange(suffixindex + 1, converted.Count - suffixindex - 1).ToArray());
                                            // check every substring for a city name
                                            foreach (var y in substrings)
                                            {
                                                // one of the substrings is a city name
                                                if (cities.Contains(Utils.RemoveDiacritics(y), StringComparer.OrdinalIgnoreCase))
                                                {
                                                    potentialcity = y;
                                                    //Console.WriteLine("potential city: " + potentialcity);
                                                }
                                            }
                                            // city not found
                                            if (potentialcity == "")
                                            {
                                                DBIO.WriteError(conString, raw);
                                                break;
                                            }
                                            else
                                            {
                                                // last iteration of potentialcity should be the city name (longest one)
                                                city = potentialcity;
                                                streetnum = converted.ElementAt(numindex);
                                                streetname = String.Join(" ", converted.GetRange(0, suffixindex + 1));
                                                cityfound = true;
                                                numfound = true;
                                                Console.WriteLine("{0, -15} {1, -25} {2, -40} {3, -10} {4, -15}", streetnum, streetname, city, province, postalcode);
                                            }
                                        }
                                    }
                                    // street number and street ending in the middle, so city is numindex + 1 onwards, and street is 0 to suffixindex
                                    else
                                    {
                                        // get all substrings from 0 to index of potential street number
                                        substrings = Utils.AllSubstrings(converted.GetRange(numindex + 1, converted.Count - 1 - numindex).ToArray());
                                        // check every substring for a city name
                                        foreach (var y in substrings)
                                        {
                                            // one of the substrings is a city name
                                            if (cities.Contains(Utils.RemoveDiacritics(y), StringComparer.OrdinalIgnoreCase))
                                            {
                                                potentialcity = y;
                                                // Console.WriteLine("potential city: " + potentialcity);
                                            }
                                        }
                                        // city not found
                                        if (potentialcity == "")
                                        {
                                            DBIO.WriteError(conString, raw);
                                            break;
                                        }
                                        else
                                        {
                                            // last iteration of potentialcity should be the city name (longest one)
                                            city = potentialcity;
                                            cityfound = true;
                                            streetnum = converted.ElementAt(numindex);
                                            numfound = true;
                                            streetname = String.Join(" ", converted.GetRange(0, suffixindex + 1));
                                            Console.WriteLine("{0, -15} {1, -25} {2, -40} {3, -10} {4, -15}", streetnum, streetname, city, province, postalcode);
                                        }
                                    }
                                }
                                potentialstreet = potentialstreet + " ";
                                pointer++;
                            }
                        }
                    }
                }
            }// province and postal code could not be extracted, save in "error" database
            else
            {
                DBIO.WriteError(conString, raw);
            }
        }
    }
}
