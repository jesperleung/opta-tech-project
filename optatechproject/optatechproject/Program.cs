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
                using (TextFieldParser parser = new TextFieldParser(@"..\..\..\..\" + filename))    // getting data file from root folder of project and provided filename
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
                            // clean up all fields by taking out anything that isn't alphanumeric or a space
                            string clean = Regex.Replace(field, "[^A-Za-z0-9 ]", "");
                            Console.WriteLine(clean);
                        }
                    }
                }
            }catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        public static void Main(string[] args)
        {

            Console.Write("Please type the filename of the input data file: ");
            string inputfilename = Console.ReadLine();
            Console.WriteLine(inputfilename);

            LoadCSV(inputfilename);

            ConnectToDB();  // connect to database

            Console.WriteLine("All done. Press any key to finish...");
            Console.ReadKey(true);
        }
    }
}
