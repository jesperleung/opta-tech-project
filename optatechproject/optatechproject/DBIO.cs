using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OptaTechProject
{
    class DBIO
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

    }
}
