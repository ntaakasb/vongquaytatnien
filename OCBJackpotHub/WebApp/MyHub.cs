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
            LuckyNumber _item = new LuckyNumber();
            string number = string.Empty;

            if (lsNumber != null && lsNumber.Any())
            {
               
               
               
                var lsResult = _services.getAllResult();
                if(lsResult != null && lsResult.Any())
                {
                    var lsNewMerge = lsNumber.Where(x => !lsResult.Where(es => es.number.ToString().Trim() == x.number1.ToString() + x.number2.ToString() + x.number3.ToString()).Any()).ToList();

                    Random r = new Random();
                    int num = r.Next(0, lsNumber.Count);
                    
                    _item = lsNewMerge[num];
                    number = _item.number1.ToString() + _item.number2.ToString() + _item.number3.ToString();
                }

                _services.writeResult(number);
                Clients.All.returnNumber(JsonConvert.SerializeObject(_item));
            }

            
        }
        public void Send(string number)
        {
            // Call the broadcastMessage method to update clients.
            Clients.All.broadcastMessage(number, DateTime.Now.ToString());
        }

        public void LoadAll()
        {
            Services _services = new Services();
            var lsResult = _services.getAllResult();
            if(lsResult != null && lsResult.Any())
            {
                Clients.All.loadAllResult(JsonConvert.SerializeObject(lsResult));

            }
          
        }

        public void ReturnList(string number)
        {
            Services _services = new Services();
            _services.removeResult(number);
        }

        public void DirectPage()
        {
            Clients.All.loadPage();
        }

    }
}