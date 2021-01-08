using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using WebApp.Model;
using System.Web.SessionState;
using System.Xml;
using System.Xml.Linq;

namespace WebApp
{
    public class Services 
    {
        public string sessionKey = "_OCB_spinner";

        public bool IsReusable => throw new NotImplementedException();

        public List<LuckyNumber> loadNumber()
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
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
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

        public List<int> readResult()
        {
            XmlTextReader textReader = new XmlTextReader("C:\\books.xml");
            textReader.Read();
            // If the node has value  
            while (textReader.Read())
            {
                // Move to fist element  
                textReader.MoveToElement();
                Console.WriteLine("XmlTextReader Properties Test");
                Console.WriteLine("===================");
                // Read this element's properties and display them on console  
                Console.WriteLine("Name:" + textReader.Name);
                Console.WriteLine("Base URI:" + textReader.BaseURI);
                Console.WriteLine("Local Name:" + textReader.LocalName);
                Console.WriteLine("Attribute Count:" + textReader.AttributeCount.ToString());
                Console.WriteLine("Depth:" + textReader.Depth.ToString());
                Console.WriteLine("Line Number:" + textReader.LineNumber.ToString());
                Console.WriteLine("Node Type:" + textReader.NodeType.ToString());
                Console.WriteLine("Attribute Count:" + textReader.Value.ToString());
            }
            return new List<int>();
        }

        public void writeResult(string number)
        {
            string url = System.Web.HttpContext.Current.Server.MapPath("/Data/Result.xml");
            XElement xEle = XElement.Load(url);
            xEle.AddFirst(new XElement("Result",
                new XElement("Number", number),
                new XElement("Time", DateTime.Now.ToString()), new XAttribute("id", number)));
            xEle.Save(url);
        }

        public void removeResult(string id)
        {
            try
            {
                string url = System.Web.HttpContext.Current.Server.MapPath("/Data/Result.xml");
                var document = XDocument.Load(url);
                var deleteQuery = document.Element("Results").Elements("Result").Where(x => x.Attribute("id").Value == id.Trim()).FirstOrDefault();
                if (deleteQuery != null)
                {
                    deleteQuery.Remove();
                    document.Save(url);
                }
            }
          catch(Exception ex)
            {

            }

        }

        public List<ResultModel> getAllResult()
        {
            string url = System.Web.HttpContext.Current.Server.MapPath("/Data/Result.xml");
            XDocument doc = XDocument.Load(url);
            List<ResultModel> lsResult = new List<ResultModel>();
            var elements = from xEle in doc.Descendants("Result") select xEle;
            foreach (XElement result in elements)
            {
                ResultModel item = new ResultModel();
                item.number = result.Element("Number").Value;
                item.CreatedDate = result.Element("Time").Value;
                lsResult.Add(item);
            }
            return lsResult;
        }
    }
}