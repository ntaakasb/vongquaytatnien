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
                LuckyNumber _item = new LuckyNumber();
                _item = lsNumber[num];
                int number = int.Parse(_item.number1.ToString() + _item.number2.ToString() + _item.number3.ToString());
                //_services.writeResult(number);
                //_services.removeResult();


                Clients.All.returnNumber(JsonConvert.SerializeObject(_item));
            }

            
        }
        public void Send(int number1, int number2, int number3)
        {
            // Call the broadcastMessage method to update clients.
            Clients.All.broadcastMessage(number1, number2, number3, DateTime.Now.ToString());
        }

        

       
    }
}