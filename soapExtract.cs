/*
Extract data from application via API and load into DB.
Utilizes SSIS and PowerShell script as well
*/

using System;
using System.Net;
using System.Xml;
using System.IO;
using System.Management.Automation;

namespace Records
{
    class Program
    {
        static void Main(string[] args)
        {
            Program obj = new Program();
            
            //takes UTC into account
            string stDate = DateTime.Now.AddDays(-1).ToString("yyyy'-'MM'-'dd")+"T04:00:00"; // gets previous day 
            string enDate = DateTime.Now.ToString("yyyy'-'MM'-'dd") + "T03:59:59"; // gets today
            //Console.WriteLine("getting data from " + stDate + " to " + enDate);
            int count = 0;
            do
            {
                try
                {
                    obj.RecordDate(stDate, enDate); // call RecordDate with data params
                    break; //continues on if successful
                }
                catch (WebException e)
                {
                    using (StreamWriter stream = File.AppendText(@"\\reportingServer\c$\soap\log.txt"))
                    {
                        string errorTime = DateTime.Now.ToString();
                        stream.WriteLine(errorTime + " : " + e.Message); //log error
                    }
                }
                finally
                {
                    count++;
                }
            }
            while (count < 5); //try 5 times to connect if there is an problem
        }
        public void RecordDate(string stDate, string enDate)
        {
            HttpWebRequest request = CreateSOAPWebRequest(); // init web request
            XmlDocument SOAPReqBody = new XmlDocument();  // create XML document object

                /* SOAP envelope message request; flip bools or enter variables here */
                SOAPReqBody.LoadXml(@"<soapenv:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sour=""http://source.myapi.com"">
   <soapenv:Header/>
   <soapenv:Body>
    SOAP request here!!
   </soapenv:Body>
</soapenv:Envelope>");

            using (Stream stream = request.GetRequestStream()) // init request stream to write incoming data
            {
                SOAPReqBody.Save(stream);
            }
            using (WebResponse Serviceres = request.GetResponse()) // returns reponse
            {
                using (StreamReader rd = new StreamReader(Serviceres.GetResponseStream())) // handles the return stream
                {
                    var ServiceResult = rd.ReadToEnd(); // read whole message
                    XmlDocument doc = new XmlDocument(); // new XML object to place message
                    doc.LoadXml(ServiceResult);
                    doc.PreserveWhitespace = true;
                    Console.WriteLine(doc.OuterXml);
                    doc.Save(@"\\reportingServer\c$\soap\Records\output\Records.xml"); // save to file
                }
                /* calls powershell script to convert to csv, load into DB, email notification when complete */
                string script = File.ReadAllText(@"\\reportingServer\c$\soap\Records\xmlToCsv.ps1"); // load script into C#
                PowerShell ps = PowerShell.Create(); // create powershell instance
                    ps.AddScript(script).Invoke(); // run powershell script
            }
        }
        /* makes the SOAP request */
        public HttpWebRequest CreateSOAPWebRequest()
        {   
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(@"http://apiserverweb01/WebInterface/services/API_Service?wsdl"); //production interface
            request.Headers.Add(@"SOAPAction:http://this.is.the.header/q");
            request.ContentType = "text/xml;charset=UTF-8";
            request.Accept = "text/xml";
            request.Method = "POST";
            return request;
        }
    }
}
