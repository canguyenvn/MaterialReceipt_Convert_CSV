using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace MaterialReceipt_Convert_CSV
{
    class CsvFile
    {
        public StreamWriter Open()
        {

            // Get DateTime Now
            DateTime dt = DateTime.Now;
            string OutPutDateTime = dt.ToString("yyyyMMddHHmmss");

            // Get CSV Output Folder at Desktop
            string CsvDir = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string CsvFile = System.Configuration.ConfigurationManager.AppSettings["CsvFileName"] + OutPutDateTime + ".csv";
            // Open CSV File
            StreamWriter FileIO = new StreamWriter(CsvDir + "\\" + CsvFile, false, Encoding.UTF8);
            FileIO.WriteLine(System.Configuration.ConfigurationManager.AppSettings["Label"]);

            return FileIO;
        }

        public void Line(StreamWriter FileIO, string CsvData)
        {
            // OutPut Line
            FileIO.WriteLine(CsvData);
        }

        public void Close(StreamWriter FileIO)
        {
            // Close CsvFile
            FileIO.Close();
        }

    }
}
