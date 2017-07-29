using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using ConsoleApplication1.Properties;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            //string connetionString = null;
            SqlConnection connection;
            SqlCommand command;
            string sql = null;
            //connetionString = @"Data Source=ls-its-db;Initial Catalog=DAL1_IDB;User ID=ADS\pmoldenhauer;Password=";
            sql = "SELECT Roadway_Name,  Is_Only_Cross_Street FROM [DAL1_IDB].[dbo].[CT_ROADWAY]";

            SqlConnectionStringBuilder csb = new SqlConnectionStringBuilder();
            csb.DataSource = "ls-its-db";
            csb.IntegratedSecurity = true;
            csb.UserID = "ADS\\pmoldenhauer";
            csb.Password = Settings.Default.SettingsKey;

            ISet<string> roadways = new HashSet<string>();
            ISet<string> crossStreets = new HashSet<string>();

            using (connection = new SqlConnection(csb.ConnectionString))
            {
                try
                {
                    connection.Open();
                    command = new SqlCommand(sql, connection);
                    using (SqlDataReader dataReader = command.ExecuteReader())
                    {
                        while (dataReader.Read())
                        {
                            //Console.WriteLine(dataReader.GetValue(0) + " - " + dataReader.GetValue(1)); 

                            if (dataReader.GetBoolean(1))
                            {
                                crossStreets.Add(dataReader.GetValue(0).ToString());
                            }
                            else
                            {
                                roadways.Add(dataReader.GetValue(0).ToString());
                            }
                        }
                    }

                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\pmoldenhauer\Desktop\Copy of RITMS roadways and cross streets.xlsx");
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    int roadway_name = 4;
                    int cross_street_name = 10;

                    // start looping at row 2 where data is - row 1 is the headings in excel
                    for (int i = 2; i <= 5; i++)
                    {
                        // Check the roadway names in Excel
                        if (xlRange.Cells[i, roadway_name] != null && xlRange.Cells[i, roadway_name].Value2 != null)
                        {
                            string roadwayName = xlRange.Cells[i, roadway_name].Value2.ToString();
                            if (!roadways.Contains(roadwayName))
                            {
                                // INSERT roadwayName into database if not in already
                                try
                                {
                                    sql =
                                        "INSERT INTO [DAL1_IDB].[dbo].[CT_ROADWAY] ([Roadway_Name],[Is_Only_Cross_Street]) VALUES ('" +
                                        roadwayName + "',0)";
                                    command = new SqlCommand(sql, connection);
                                    int result = command.ExecuteNonQuery();
                                    Console.WriteLine("Roadway inserted into database: " + roadwayName);

                                    command.Dispose();
                                }
                                catch (SqlException ex)
                                {
                                    Console.WriteLine("ERROR! Insert failed!");
                                }
                            }
                            
                            roadways.Add(roadwayName);
                        }
                        
                        // Check the cross street names in Excel
                        if (xlRange.Cells[i, cross_street_name] != null && xlRange.Cells[i, cross_street_name].Value2 != null)
                        {
                            string roadwayName = xlRange.Cells[i, cross_street_name].Value2.ToString();
                            if (!crossStreets.Contains(roadwayName))
                            {
                                // INSERT cross street into database if not in already
                                try
                                {
                                    sql =
                                        "INSERT INTO [DAL1_IDB].[dbo].[CT_ROADWAY] ([Roadway_Name],[Is_Only_Cross_Street]) VALUES ('" +
                                        roadwayName + "',1)";
                                    command = new SqlCommand(sql, connection);
                                    int result = command.ExecuteNonQuery();
                                    Console.WriteLine("Cross Street inserted into database: " + roadwayName);

                                    command.Dispose();
                                }
                                catch (SqlException ex)
                                {
                                    Console.WriteLine("ERROR! Insert failed!");
                                }
                            }
                            
                            crossStreets.Add(roadwayName);
                        }
                    }

                    //cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    //release com objects to fully kill excel process from running in the background
                    Marshal.ReleaseComObject(xlRange);
                    Marshal.ReleaseComObject(xlWorksheet);

                    //close and release
                    xlWorkbook.Close();
                    Marshal.ReleaseComObject(xlWorkbook);

                    //quit and release
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);

                    command.Dispose();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Can not open connection! ");
                }
            }
            Console.ReadKey();
        }
    }
}
