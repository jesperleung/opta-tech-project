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
        public static string ConnectToDB()
        {
            // connection string from Server Explorer in VS after adding database
            string conString = "Data Source=.\\SQLEXPRESS;Integrated Security=True";
            try
            {
                Console.Write("Connecting to SQL Server ... ");
                using (SqlConnection con = new SqlConnection(conString))
                {
                    con.Open();

                    string sql = "DROP DATABASE IF EXISTS optatechproject; CREATE DATABASE optatechproject; DROP TABLE IF EXISTS dbo.rawdata; DROP TABLE IF EXISTS dbo.complete; DROP TABLE IF EXISTS dbo.error;";

                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        // drop old database and tables if they exist
                        cmd.ExecuteNonQuery();

                    }

                    // create new database and tables
                    sql = "USE optatechproject; CREATE TABLE dbo.rawdata (ID int PRIMARY KEY NOT NULL, rawtext varchar(255) NOT NULL);";
                    sql = sql + "CREATE TABLE dbo.complete (ID int PRIMARY KEY NOT NULL, streetnum varchar(255) NOT NULL, streetname varchar(255) NOT NULL, city varchar(255) NOT NULL, province varchar(2) NOT NULL, postalcode varchar(6) NOT NULL);";
                    sql = sql + "CREATE TABLE dbo.error (ID int PRIMARY KEY NOT NULL, string varchar(255) NOT NULL);";

                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        cmd.ExecuteNonQuery();
                    }



                    Console.WriteLine("Ready.");

                }
            }
            catch (SqlException e)
            {
                // if error, print error to console
                Console.WriteLine(e.ToString());
            }
            return conString;
        }
        // writes parsed address to database table dbo.complete
        public static void WriteComplete(string conString, string streetnum, string streetname, string city, string province, string postalcode)
        {
            int max = 0;

            try
            {
                using (SqlConnection con = new SqlConnection(conString))
                {
                    con.Open();

                    String sql = "USE optatechproject; SELECT MAX(ID) FROM dbo.complete;";
                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                try
                                {
                                    if (reader.IsDBNull(0))
                                    {
                                        max = 0;
                                    }
                                    else
                                    {
                                        max = reader.GetInt32(0) + 1;
                                    }
                                }
                                catch (Exception e)
                                {
                                    max = 0;
                                }
                            }
                        }
                    }
                    sql = "USE optatechproject; INSERT dbo.complete(ID, streetnum, streetname, city, province, postalcode)" +
                    "VALUES (@ID, @streetnum, @streetname, @city, @province, @postalcode);";

                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        cmd.Parameters.AddWithValue("@ID", max);
                        cmd.Parameters.AddWithValue("@streetnum", streetnum);
                        cmd.Parameters.AddWithValue("@streetname", streetname);
                        cmd.Parameters.AddWithValue("@city", city);
                        cmd.Parameters.AddWithValue("@province", province);
                        cmd.Parameters.AddWithValue("@postalcode", postalcode);

                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (SqlException e)
            {
                Console.WriteLine(e.ToString());
            }
        }
        //
        // writes error address to database table dbo.error
        public static void WriteError(string conString, string raw)
        {
            int max = 0;

            try
            {
                using (SqlConnection con = new SqlConnection(conString))
                {
                    con.Open();

                    String sql = "USE optatechproject; SELECT MAX(ID) FROM dbo.error;";
                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                try
                                {
                                    if (reader.IsDBNull(0))
                                    {
                                        max = 0;
                                    }
                                    else
                                    {
                                        max = reader.GetInt32(0) + 1;
                                    }
                                }
                                catch (Exception e)
                                {
                                    max = 0;
                                }
                            }
                        }
                    }
                    sql = "USE optatechproject; INSERT dbo.error(ID, string)" +
                        "VALUES (@ID, @string);";

                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        cmd.Parameters.AddWithValue("@ID", max);
                        cmd.Parameters.AddWithValue("@string", raw);

                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (SqlException e)
            {
                Console.WriteLine(e.ToString());
            }
        }
        // writes all raw entries into dbo.raw
        public static void WriteRaw(string conString, string raw)
        {
            int max = 0;

            try
            {
                using (SqlConnection con = new SqlConnection(conString))
                {
                    con.Open();

                    String sql = "USE optatechproject; SELECT MAX(ID) FROM dbo.rawdata;";
                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                try
                                {
                                    if (reader.IsDBNull(0))
                                    {
                                        max = 0;
                                    }
                                    else
                                    {
                                        max = reader.GetInt32(0) + 1;
                                    }
                                }
                                catch (Exception e)
                                {
                                    max = 0;
                                }
                            }
                        }
                    }
                    sql = "USE optatechproject; INSERT dbo.rawdata(ID, rawtext)" +
                        "VALUES (@ID, @rawtext);";

                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        cmd.Parameters.AddWithValue("@ID", max);
                        cmd.Parameters.AddWithValue("@rawtext", raw);

                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (SqlException e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}
