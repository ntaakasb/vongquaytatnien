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
            Services _services = new Services();
            List<LuckyNumber> lsNumber = new List<LuckyNumber>();
            lsNumber = _services.loadNumber();


            if (lsNumber != null && lsNumber.Any())
            {
                Random r = new Random();
                int num = r.Next(0, lsNumber.Count);
            }

            Clients.All.returnNumber(JsonConvert.SerializeObject(lsNumber));
        }
        public void Send(string name, string message)
        {
            // Call the broadcastMessage method to update clients.
            Clients.All.broadcastMessage(name, message);
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