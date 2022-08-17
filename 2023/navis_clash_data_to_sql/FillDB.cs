using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace fill_in_db
{
    public class FillDB
    {
        ////Old code to call python script to write to sqlite database file--------------------------
        //public static void Main_P()
        //{
        //    try
        //    {
        //        string exe_path = @"C:\Program Files\Autodesk\IPS Navis clashes to database\clashes_to_sql_v1.0.0.exe";
        //        //string exe_path = @"C:\Program Files\Autodesk\Navisworks Manage 2021\Plugins\navis_clash_data_to_sql\clashes_to_sql_v1.0.0.exe";
        //        Process.Start(exe_path);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error accessing clashes exe file", ex.Message);
        //    }
        //}
        ////-------------------------------------------------------------------------------------------
        

        ////Code used to test connection to database-----------------------------------------------------
        public static string SQL_test()
        {
            String cs = "Data Source = USBLB1DB002\\APP05;Initial Catalog=NWClashData;Integrated Security=true";
            try
            {
                SqlConnection conn = new SqlConnection(cs);
                conn.Open();

                List<string> Table_list = new List<string>();
                using (SqlCommand com = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES", conn))
                {
                    using (SqlDataReader reader = com.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Table_list.Add((string)reader["TABLE_NAME"]);
                        }
                    }
                }
                conn.Close();
                foreach (string sTable in Table_list)
                {
                    //MessageBox.Show("Table name", sTable);
                    return "Success";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error connecting to datbase. Check network/VPN Connection.", ex.Message);
                return "Fail";
            }
            return "Fail";
        }
        ////--------------------------------------------------------------------------------------

        public static string Main_C(DataTable input_table)
        {
            String cs = "Data Source = USBLB1DB002\\APP05;Initial Catalog=NWClashData;Integrated Security=true";
            //String cs = "Data Source = USBLB1DB002\\APP05;Initial Catalog=NWClashDataTest;Integrated Security=true";
            try
            {
                SqlConnection conn = new SqlConnection(cs);
                conn.Open();
                using (SqlBulkCopy bulkCopyTable = new SqlBulkCopy(conn))
                {
                foreach (DataColumn c in input_table.Columns)
                    bulkCopyTable.ColumnMappings.Add(c.ColumnName, c.ColumnName);
                    bulkCopyTable.DestinationTableName = "clash_results";
                    int numberOfRecords = input_table.Rows.Count;
                    if (numberOfRecords > 0)
                    {
                        try
                        {
                            bulkCopyTable.WriteToServer(input_table);
                            conn.Close();
                            //Pass
                            string testName = "P," + input_table.TableName.ToString();
                            return testName;
                            //MessageBox.Show("Data was sucessfully added to the database for table " + input_table.TableName);
                        }
                        catch
                        {
                            //MessageBox.Show(ex.Message, "Error writing" + input_table.TableName + "data to database.");
                            conn.Close();
                            //Fail
                            string testName = "F," + input_table.TableName.ToString();
                            return testName;
                        }
                    }
                    else
                    {
                        //Empty test
                        string itemResult = "E," + input_table.TableName.ToString();
                        return itemResult;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error with connecting to database.");
                string itemResult = "Connection Failed";
                return itemResult;
            }
        }
    }
}
