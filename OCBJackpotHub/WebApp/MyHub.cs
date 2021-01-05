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
using Newtonsoft.Json;
using OfficeOpenXml;
using WebApp.Model;

namespace WebApp
{
    [HubName("myHub")]
    public class MyHub : Hub
    {
        public void LoadNumber()
        {
            List<LuckyNumber>  lsNumber = loadNumber();
            Clients.All.returnNumber(JsonConvert.SerializeObject(lsNumber));
        }
        public void Send(string name, string message)
        {
            // Call the broadcastMessage method to update clients.
            Clients.All.broadcastMessage(name, message);
        }

        private List<LuckyNumber> loadNumber()
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
                DataTable dt = ds.Tables[0];
                List<LuckyNumber> luckyModel = new List<LuckyNumber>();
                if(dt != null)
                {
                   foreach(DataRow dr in dt.Rows)
                    {

                        LuckyNumber item = new LuckyNumber();
                        item.index = dr[0] != null ? int.Parse(dr[0].ToString()) : 0;
                        item.number1 = dr[1] != null ? int.Parse(dr[1].ToString()) : 0;
                        item.number2 = dr[2] != null ? int.Parse(dr[2].ToString()) : 0;
                        item.number3 = dr[3] != null ? int.Parse(dr[3].ToString()) : 0;
                        item.displayed = dr[4] != null ? int.Parse(dr[4].ToString()) : 0;
                        item.prizeCode = dr[5] != null ? int.Parse(dr[5].ToString()) : 0;
                        luckyModel.Add(item);
                    }
                }
                return luckyModel;
            }
            catch (Exception ex)
            {
                string e = ex.Message;
                return new List<LuckyNumber>();
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
            string path = @"D:\WorkSpace\Project Outsource\slotmachine\vongquaytatnien\OCBJackpotHub\WebApp\dataNumber.xlsx";
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