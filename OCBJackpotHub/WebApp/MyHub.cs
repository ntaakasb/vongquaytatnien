using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using Microsoft.AspNet.SignalR;
using Microsoft.AspNet.SignalR.Hubs;
using WebApp.Data;

namespace WebApp
{
    [HubName("myHub")]
    public class MyHub : Hub
    {
        public void Hello()
        {
            readExcel();
            Clients.All.hello();
        }
        public void Send(string name, string message)
        {
            // Call the broadcastMessage method to update clients.
            Clients.All.broadcastMessage(name, message);
        }

        public void readExcel()
        {
            string con =
    @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\WorkSpace\Project Outsource\slotmachine\vongquaytatnien\Data.xlsx;" +
    @"Extended Properties='Excel 12.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand cmd = new OleDbCommand("select * from [Sheet1$]", connection);
                //Or Use OleDbCommand
                OleDbDataAdapter oleda = new OleDbDataAdapter();

                oleda.SelectCommand = cmd;

                // Create a DataSet which will hold the data extracted from the worksheet.
                DataSet ds = new DataSet();

                // Fill the DataSet from the data extracted from the worksheet.
                oleda.Fill(ds, "Employees");
            }
        }
    }
}