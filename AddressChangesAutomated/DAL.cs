﻿using System;
using System.Data;
using System.Data.SqlClient;

namespace AddressChangesAutomated
{
    class DAL
    {

        public static bool RunNonQuery(string query, string uname, string pwd)
        {
            string Connstring = "Data Source =tvco360sql01; Initial Catalog = tvcOperations ; User Id = " + uname + "; Password =" + pwd + ";Connection Timeout=1000; Application Name=SpreadsheetImporter";

            bool Result = false;
            try
            {
                using (SqlConnection connection = new SqlConnection(Connstring))
                {
                    SqlCommand command = new SqlCommand(query, connection);
                    command.CommandTimeout = 100;
                    command.Connection.Open();
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

                throw;
            }


            return Result;
        }


        public static DataTable GetDataTable(string query, string uname, string pwd)
        {
            string Connstring = "Data Source =tvco360sql01; Initial Catalog = tvcOperations ; User Id = " + uname + "; Password =" + pwd + ";Connection Timeout=1000; Application Name=SpreadsheetImporter";
            try
            {
                using (SqlConnection con = new SqlConnection(Connstring))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandText = query;
                        cmd.CommandTimeout = 300000;
                        using (SqlDataAdapter sda = new SqlDataAdapter())
                        {
                            cmd.Connection = con;
                            sda.SelectCommand = cmd;

                            using (DataSet ds = new DataSet())
                            {
                                DataTable dt = new DataTable();
                                sda.Fill(dt);
                                return dt;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var dt = new DataTable
                {
                    Columns = { { "Message", typeof(string) } }
                };
                if (ex.Message.ToUpper().Contains("LOGIN FAILED FOR"))
                {
                    dt.Rows.Add("THE LOGIN FAILED");
                }
                else
                {
                    throw;
                }
                return dt;

                throw;
            }
        }
    }
}
