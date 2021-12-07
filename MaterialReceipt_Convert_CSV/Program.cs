using System;
using System.IO;
using System.Text;
using System.Linq;
//using System.​Windows.Forms;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using ClosedXML.Excel;
using static System.Console;
// Use Postgresql
using Npgsql;

namespace MaterialReceipt_Convert_CSV
{
    class Program
    {
        static void Main(string[] args)
        {
            string CsvData;
            int LineNo = 10;

            //MessageBox.Show("Please enter the correct value.", MessageBoxButtons.OK, MessageBoxIcon.Error);

            // Get Connecting infomation for LEAP DataBase 
            string LEAPDataBase = System.Configuration.ConfigurationManager.AppSettings["LeapDataBase"];
            // Get Csv Column Infomation
            int StartRow = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["StartRow"]);
            int InboundDateColumn = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["InboundDateColumn"]);
            int ItemCodeColumn = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["ItemCodeColumn"]);
            int DiscriptionColumn = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["DiscriptionColumn"]);
            int InVoiceDateColumn = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["InVoiceDateColumn"]);
            int InVoiceNoColumn = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["InVoiceNoColumn"]);
            int UnitColumn = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["UnitColumn"]);
            int InboundQTYColumn = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["InboundQTYColumn"]);
            int PurchaseOrderColumn = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["PurchaseOrderColumn"]);
            int PurchaseOrderLineColumn = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["PurchaseOrderLineColumn"]);
            int SupplierColumn = Int32.Parse(System.Configuration.ConfigurationManager.AppSettings["SupplierColumn"]);

            // Open CSV File
            CsvFile OutPutCsv = new CsvFile();
            StreamWriter FileIO = OutPutCsv.Open();

            //Display command line arguments
            Console.WriteLine(System.Environment.CommandLine);

            // Get command line paramater for Drag & Drop 
            string[] cmds = System.Environment.GetCommandLineArgs();
            string ExcelFilePath = cmds[1];

            var ConnDB = new NpgsqlConnection(LEAPDataBase);
            {
                // Connect to LEAD DataBase (PostgreSQL) 
                ConnDB.Open();
                Console.WriteLine("Connection success!");
                {
                    try
                    {

                        // Open MaterialReceiptFile
                        var workbook = new XLWorkbook(ExcelFilePath);
                        // Read INPUT WorkSheet
                        var worksheet = workbook.Worksheet("INPUT");
                        if (worksheet is null)
                        {
                            // Display Alert
                            Console.WriteLine("%%% Not Exist input WorkSheet");
                        }

                        var range = worksheet.RangeUsed();
                        int RowCount = range.RowCount();
                        //int ColCount = range.ColumnCount();
                        for (int Row = StartRow; Row <= RowCount; Row++)
                        {
                            string PurchaseOrder = range.Cell(Row, PurchaseOrderColumn).GetFormattedString().Trim();

                            if (PurchaseOrder != "")
                            {

                                //MessageBox.Show("Purcase Order=" + PurcaseOrder, "Comfirm", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                Console.Write("{0}\t", PurchaseOrder);

                                string query = "select warehouse from material where purchase_order='" + PurchaseOrder + "'";
                                //string query = "select warehouse from material";
                                var SqlCmd = new NpgsqlCommand(query, ConnDB);

                                var Reader = SqlCmd.ExecuteReader();
                                Reader.Read();
                                //while (Reader.Read())
                                //{
                                string WareHouse = Reader.GetString(0);
                                Reader.Close();
                                string Location = System.Configuration.ConfigurationManager.AppSettings[WareHouse];
                                Console.WriteLine("value : {0} {1}", WareHouse, Location);

                                // Setting CsvData
                                CsvData = System.Configuration.ConfigurationManager.AppSettings["AD_Org_ID"];
                                CsvData += "," + System.Configuration.ConfigurationManager.AppSettings["C_DocType_ID"];
                                DateTime Dt = range.Cell(Row, InboundDateColumn).GetDateTime();
                                CsvData += "," + Dt.ToString("yyyy/MM/dd");
                                CsvData += "," + Dt.ToString("yyyy/MM/dd");
                                CsvData += "," + range.Cell(Row, SupplierColumn).GetString().Trim();
                                CsvData += "," + WareHouse;
                                CsvData += ",";
                                CsvData += "," + range.Cell(Row, PurchaseOrderColumn).GetString().Trim();
                                CsvData += "," + LineNo;
                                CsvData += "," + range.Cell(Row, PurchaseOrderColumn).GetString().Trim();
                                CsvData += "," + range.Cell(Row, PurchaseOrderLineColumn).GetString().Trim();
                                CsvData += "," + range.Cell(Row, ItemCodeColumn).GetString().Trim();
                                CsvData += "," + range.Cell(Row, InboundQTYColumn).GetString().Trim();
                                CsvData += "," + range.Cell(Row, UnitColumn).GetString().Trim();
                                CsvData += "," + Location;
                                CsvData += "," + range.Cell(Row, DiscriptionColumn).GetString().Trim();
                                CsvData += "," + range.Cell(Row, InVoiceNoColumn).GetString().Trim();
                                CsvData += "," + range.Cell(Row, InVoiceDateColumn).GetString().Trim();
                                // OutPut Csv File
                                OutPutCsv.Line(FileIO, CsvData);

                                // SerialNo CountUp
                                LineNo += 10;
                            }
                        }
                    }
                    catch (Npgsql.PostgresException ex)
                    {
                        Console.WriteLine(ex.SqlState);
                    }

                }
                // Close DataBase
                ConnDB.Close();
            }
            // range.Dispose();

            // Close CsvFile
            OutPutCsv.Close(FileIO);
        }
    }
}

