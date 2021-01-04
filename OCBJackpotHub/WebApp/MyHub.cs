using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using Microsoft.AspNet.SignalR;
using Microsoft.AspNet.SignalR.Hubs;
using OfficeOpenXml;

namespace WebApp
{
    [HubName("myHub")]
    public class MyHub : Hub
    {
        public void Hello()
        {
            editRow();
            Clients.All.hello();
        }
        public void Send(string name, string message)
        {
            // Call the broadcastMessage method to update clients.
            Clients.All.broadcastMessage(name, message);
        }

        public void readExcel()
        {
            string connString = ConfigurationManager.ConnectionStrings["xlsx"].ConnectionString;
            // Create the connection object
            OleDbConnection oledbConn = new OleDbConnection(connString);
            try
            {
                // Open connection
                oledbConn.Open();

                // Create OleDbCommand object and select data from worksheet Sheet1
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$]", oledbConn);

                // Create new OleDbDataAdapter
                OleDbDataAdapter oleda = new OleDbDataAdapter();

                oleda.SelectCommand = cmd;

                // Create a DataSet which will hold the data extracted from the worksheet.
                DataSet ds = new DataSet();

                // Fill the DataSet from the data extracted from the worksheet.
                oleda.Fill(ds, "Employees");

                // Bind the data to the GridView
            }
            catch (Exception ex)
            {
                string e = ex.Message;
            }
            finally
            {
                // Close connection
                oledbConn.Close();
            }
        }

        public void editRow()
        {
        
            // path to your excel file
            string path = @"F:\workspace\git\vongquaytatnien\OCBJackpotHub\WebApp\datanumber.xlsx";
            FileInfo fileInfo = new FileInfo(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            // get number of rows in the sheet
            int rows = worksheet.Dimension.Rows; // 10

            // loop through the worksheet rows
            for (int i = 1; i <= rows; i++)
            {
                // replace occurences
                worksheet.Cells[i, 2].Value = "ABCD";
            }

            // save changes
            package.Save();
        }
    }
}