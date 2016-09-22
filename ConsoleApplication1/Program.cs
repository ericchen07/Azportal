using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

using System.IO;
using System.Net;

using Newtonsoft.Json.Linq;

using System.Net.Mail;

using System.Data.SqlClient;

namespace Microsoft.SDK.SharePointServices.Samples
{
    class RetrieveAllListProperties
    {
        static void Main(string[] args)
        {
            //LSIupdate();
            //UpdateStatus();
            //GetListProperities();
            //EmailLogic();
            //TestEmail();
            ImportInfor1();
        }
        public static void LSIupdate()
        {
            string siteUrl = "http://azportal";
            using (SPSite site = new SPSite(siteUrl))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPList list = web.GetList("http://azportal/Lists/Azure Outage List");
                    SPList LSIList = web.GetList("http://azportal/Lists/Outage Count");
                    // Get the item collections which the case details haven't been updated
                    SPQuery LSIquery = new SPQuery();
                    LSIquery.Query =
                        "<Where>" +
                         "<Eq><FieldRef Name=\"Same_x0020_Active\" /><Value Type=\"Boolean\">" + "0" + "</Value></Eq>" +
                        "</Where>";
                    SPListItemCollection LSIitems = LSIList.GetItems(LSIquery);
                    foreach (SPListItem item in LSIitems) // Item in Outage Count list
                    {
                        string LSIID = (string)item["LSI ID"];
                        SPQuery fullQuery = new SPQuery();
                        // Obtain the cases under the specfic lsi
                        // Order the query result by Calling Country
                        fullQuery.Query = "<Where><Eq><FieldRef Name=\"LSI_x0020_ID\" /><Value Type=\"Text\">" + LSIID + "</Value></Eq></Where><OrderBy><FieldRef Name=\"Calling_x0020_Country\" Ascending='TRUE'></FieldRef></OrderBy>";
                        SPListItemCollection fullItems = list.GetItems(fullQuery);
                        foreach (SPListItem caseitem in fullItems) // Item in Azure Outage List
                        {
                            caseitem["LSI Active"]= item["LSI Active"];
                            caseitem.Update();
                        }
                        item["LSI Previous Active"] = item["LSI Active"];
                        item["Same Active"] = 1;
                        item.Update();
                    }
                    //RCA
                    // Get the item collections which the case details haven't been updated
                    SPQuery RCALSIquery = new SPQuery();
                    RCALSIquery.Query =
                        "<Where>" +
                         "<Eq><FieldRef Name=\"Same_x0020_RCA\" /><Value Type=\"Boolean\">" + "0" + "</Value></Eq>" +
                        "</Where>";
                    SPListItemCollection RCALSIitems = LSIList.GetItems(RCALSIquery);
                    foreach (SPListItem item in RCALSIitems) // Item in Outage Count list
                    {
                        string LSIID = (string)item["LSI ID"];
                        SPQuery fullQuery = new SPQuery();
                        // Obtain the cases under the specfic lsi
                        // Order the query result by Calling Country
                        fullQuery.Query = "<Where><Eq><FieldRef Name=\"LSI_x0020_ID\" /><Value Type=\"Text\">" + LSIID + "</Value></Eq></Where><OrderBy><FieldRef Name=\"Calling_x0020_Country\" Ascending='TRUE'></FieldRef></OrderBy>";
                        SPListItemCollection fullItems = list.GetItems(fullQuery);
                        foreach (SPListItem caseitem in fullItems) // Item in Azure Outage List
                        {
                            caseitem["LSI RCA"] = item["LSI RCA"];
                            caseitem.Update();
                        }
                        item["LSI Previous RCA"] = item["LSI RCA"];
                        item["Same RCA"] = 1;
                        item.Update();
                    }
                }
            }
        }

        public static void UpdateStatus()
        {
            string siteUrl = "http://azportal";
            using (SPSite site = new SPSite(siteUrl))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPList list = web.GetList("http://azportal/Lists/Azure Outage List");

                    // Get the item collections which the case details haven't been updated
                    SPQuery query = new SPQuery();
                    query.Query =
                        "<Where>" +
                        "<And>" +
                        "<Eq><FieldRef Name=\"LSI_x0020_Active\" /><Value Type=\"Boolean\">" + "1" + "</Value></Eq>" +
                        "<Eq><FieldRef Name=\"Active\" /><Value Type=\"Text\">" + "0" + "</Value></Eq>" +
                        "</And>" +
                        "</Where>";
                    SPListItemCollection items = list.GetItems(query);
                    // Get info and fill in all items in the item collection
                    foreach (SPListItem item in items)
                    {
                        if ((string)item["Case ID"] == null)
                            continue;

                        string SRNumber = (item["Case ID"]).ToString();
                        string MSSolveBaseURL = @"https://mssolveweb.partners.extranet.microsoft.com/MSSolveWeb/Home";

                        try
                        {
                            // call the MSSolve Web Service to get the response.
                            WebRequest webRequest = WebRequest.Create(MSSolveBaseURL + "/GetSR/" + SRNumber + "/0");
                            webRequest.Timeout = 180000;
                            ((HttpWebRequest)webRequest).UserAgent = @"Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko";
                            webRequest.UseDefaultCredentials = false;
                            webRequest.Credentials = CredentialCache.DefaultCredentials; // Use t-zefu credential
                            string text = string.Empty;
                            WebResponse response = webRequest.GetResponse();
                            Stream responseStream = response.GetResponseStream();
                            StreamReader streamReader = new StreamReader(responseStream);
                            text = streamReader.ReadToEnd();
                            response.Dispose();
                            responseStream.Close();
                            responseStream.Dispose();

                            // Use Newtonsoft to get the target data from json
                            JObject obj = JObject.Parse(text);
                            JObject ServiceRequest = (JObject)obj["Data"]["ServiceRequestResponseData"]["ServiceRequest"];
                            JObject EmailContact = (JObject)obj["Data"]["ServiceRequestResponseData"]["Contacts"];

                            // Target data are in ServiceRequest except email
                            string CustomerName = (string)ServiceRequest["CurrentAuthorizedContactIdName"];
                            string CustomerCompanyName = (string)ServiceRequest["AccountIdName"];
                            string TAMName = (string)ServiceRequest["PrimaryAccountManagerIdName"];
                            IList<JToken> emailList = EmailContact["SRContacts"].Children().ToList();
                            string CustomerContactEmail = (string)emailList[0]["PrimaryEmail"];
                            string CallingCountry = (string)ServiceRequest["CallingCountryCode"];
                            string ContractCountry = (string)ServiceRequest["ContractCountryIdName"];
                            string ServiceLevel = (string)ServiceRequest["ServiceLevelName"];

                            // Fill out the list item with the data
                            item["Case Status"] = (string)ServiceRequest["StatusName"];
                            item["Customer Company Name"] = (string)ServiceRequest["AccountIdName"];
                            item["Customer Name"] = (string)ServiceRequest["CurrentAuthorizedContactIdName"];
                            item["Customer Contact Email"] = (string)emailList[0]["PrimaryEmail"];
                            item["Premier/BC"] = (string)ServiceRequest["ServiceLevelName"];
                            item["Owner Name"] = (string)ServiceRequest["OwnerUserName"];
                            item["TAM Name"] = (string)ServiceRequest["PrimaryAccountManagerIdName"];
                            item["Calling Country"] = (string)ServiceRequest["CallingCountryCode"];
                            item["Contract Country"] = (string)ServiceRequest["ContractCountryIdName"];
                            item["Active"] = "0";

                            if (item != null)
                            {
                                item.Update();
                            }
                        }
                        catch
                        {
                            //errorMessage is not used, just in case of invalid case ID
                            string errorMessage = "case Id do not exist! or something else led to failure of getting other info";
                        }
                    }
                }
            }
        }

        public static void GetListProperities()
        {
            string siteUrl = "http://azportal";
            using (SPSite site = new SPSite(siteUrl))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPList list = web.GetList("http://azportal/Lists/Azure Outage List");

                    // Get the item collections which the case details haven't been updated
                    SPQuery query = new SPQuery();
                    query.Query = "<Where><Neq><FieldRef Name=\"Active\" /><Value Type=\"Text\">" + "0" + "</Value></Neq></Where>";// active!=0
                    SPListItemCollection items = list.GetItems(query);
                    string itemInternalName = items[0].Fields["LSI ID"].InternalName;
                    // Get info and fill in all items in the item collection
                    foreach (SPListItem item in items)
                    {
                        if ((string)item["Case ID"] == null)
                            continue;

                        string SRNumber = (item["Case ID"]).ToString();
                        string MSSolveBaseURL = @"https://mssolveweb.partners.extranet.microsoft.com/MSSolveWeb/Home";

                        try
                        {
                            // call the MSSolve Web Service to get the response.
                            WebRequest webRequest = WebRequest.Create(MSSolveBaseURL + "/GetSR/" + SRNumber + "/0");
                            webRequest.Timeout = 180000;
                            ((HttpWebRequest)webRequest).UserAgent = @"Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko";
                            webRequest.UseDefaultCredentials = false;
                            webRequest.Credentials = CredentialCache.DefaultCredentials; // Use t-zefu credential
                            string text = string.Empty;
                            WebResponse response = webRequest.GetResponse();
                            Stream responseStream = response.GetResponseStream();
                            StreamReader streamReader = new StreamReader(responseStream);
                            text = streamReader.ReadToEnd();
                            response.Dispose();
                            responseStream.Close();
                            responseStream.Dispose();

                            // Use Newtonsoft to get the target data from json
                            JObject obj = JObject.Parse(text);
                            JObject ServiceRequest = (JObject)obj["Data"]["ServiceRequestResponseData"]["ServiceRequest"];
                            JObject EmailContact = (JObject)obj["Data"]["ServiceRequestResponseData"]["Contacts"];

                            // Target data are in ServiceRequest except email
                            string CustomerName = (string)ServiceRequest["CurrentAuthorizedContactIdName"];
                            string CustomerCompanyName = (string)ServiceRequest["AccountIdName"];
                            string TAMName = (string)ServiceRequest["PrimaryAccountManagerIdName"];
                            IList<JToken> emailList = EmailContact["SRContacts"].Children().ToList();
                            string CustomerContactEmail = (string)emailList[0]["PrimaryEmail"];
                            string CallingCountry = (string)ServiceRequest["CallingCountryCode"];
                            string ContractCountry = (string)ServiceRequest["ContractCountryIdName"];
                            string ServiceLevel = (string)ServiceRequest["ServiceLevelName"];

                            // Fill out the list item with the data
                            item["Case Status"] = (string)ServiceRequest["StatusName"];
                            item["Customer Company Name"] = (string)ServiceRequest["AccountIdName"];
                            item["Customer Name"] = (string)ServiceRequest["CurrentAuthorizedContactIdName"];
                            item["Customer Contact Email"] = (string)emailList[0]["PrimaryEmail"];
                            item["Premier/BC"] = (string)ServiceRequest["ServiceLevelName"];


                            item["Owner Name"] = (string)ServiceRequest["OwnerUserName"];
                            item["TAM Name"] = (string)ServiceRequest["PrimaryAccountManagerIdName"];
                            item["Calling Country"] = (string)ServiceRequest["CallingCountryCode"];
                            item["Contract Country"] = (string)ServiceRequest["ContractCountryIdName"];
                            string email = (string)ServiceRequest["OwnerUserInternalEmail"];
                            char delimiterChars = '@';
                            string[] words = email.Split(delimiterChars);
                            string alias=words[0];
                            string queryStr = "select distinct EmployeeEmail, workgroup from vwDimEmployee where iscurrent = 'yes' and EmployeeEmail = '" + alias + "'";
                            using (SqlConnection con = new SqlConnection(@"Server=detego-ctssql;Database=ssCTSDataMart;Integrated Security=True;"))
                            {
                                using (SqlCommand cmd = new SqlCommand(queryStr, con))
                                {
                                    cmd.CommandTimeout = 2000;
                                    con.Open();
                                    using (SqlDataReader reader = cmd.ExecuteReader())
                                    {
                                        while (reader.Read())
                                        {
                                            string fields = reader["workgroup"].ToString();
                                            item["Engineer Working Group"] = fields;
                                        }
                                    }
                                }
                                
                                
                            }
                            item["Active"] = "0";

                            if (item != null)
                            {
                                item.Update();
                            }
                        }
                        catch
                        {
                            //errorMessage is not used, just in case of invalid case ID
                            string errorMessage = "case Id do not exist! or something else led to failure of getting other info";
                        }
                    }
                }
            }
        }
        public static void EmailLogic()
        {
            string siteUrl = "http://azportal";
            using (SPSite site = new SPSite(siteUrl))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPList list = web.GetList("http://azportal/Lists/Outage Count"); // The list holding count of every lsi
                    SPList fullList = web.GetList("http://azportal/Lists/Azure Outage List"); // The full list to get detailed information
                    // Get the item collections which the case details haven't been updated
                    SPQuery query = new SPQuery();
                    // IsEmailSent==Not and Amount >= 5
                    query.Query =
                        "<Where>" +
                        "<And>" +
                        "<Eq><FieldRef Name=\"IsEmailSent\" /><Value Type=\"Boolean\">" + "0" + "</Value></Eq>" +
                        "<Geq><FieldRef Name=\"Amount\" /><Value Type=\"Number\">" + "5" + "</Value></Geq>" +
                        "</And>" +
                        "</Where>";
                    SPListItemCollection items = list.GetItems(query);
                    foreach (SPListItem item in items) // Item in Outage Count list
                    {
                        string LSIID = (string)item["LSI ID"];
                        SPQuery fullQuery = new SPQuery();
                        // Obtain the cases under the specfic lsi
                        // Order the query result by Calling Country
                        fullQuery.Query = "<Where><Eq><FieldRef Name=\"Title\" /><Value Type=\"Text\">" + LSIID + "</Value></Eq></Where><OrderBy><FieldRef Name=\"Calling_x0020_Country\" Ascending='TRUE'></FieldRef></OrderBy>";
                        SPListItemCollection fullItems = fullList.GetItems(fullQuery);
                        int fullCount = fullItems.Count;
                        // Send the email
                        EmailContent(LSIID, fullCount, fullItems);
                        item["Is Email Sent"] = true;
                        item.Update();
                    }
                }
            }
        }
        public static void EmailContent(string LSINumber, int fullCount, SPListItemCollection fullItems)
        {
            //string to = "marlonj@microsoft.com";
            string to = "yanqwa@microsoft.com";
            string subject = "Azpotral outage alert (LSI ID= "+LSINumber+") Massive cases in one outage.";
            string body = "Welcome to AzPortal.<br/>There are "+ fullCount + " cases under one LSI: "+ LSINumber+
                "<br/>For detailed LSI description, please go to http://iridias/reporting/incidentlookup/"+ LSINumber+
                "<br/><br/>";
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat(body);
            // Case info table
            sb.Append("<table border=\"1\">");
            sb.AppendFormat("<tr><td>Case ID</td>"+ "<td>Customer Company Name</td>" + 
                "<td>Customer Name</td>"+ "<td>Customer Contact Email</td>" +
                "<td>Premier/BC</td>" + "<td>TAM Name</td>" +
                "<td>Calling Country</td>" + "<td>Contract Country</td>" + "</tr>");
            foreach (SPListItem fullItem in fullItems) // item in Azure Outage list
            {
                sb.AppendFormat("<tr><td>{0}</td>" + "<td>{1}</td>" +
                "<td>{2}</td>" + "<td>{3}</td>" +
                "<td>{4}</td>" + "<td>{5}</td>" +
                "<td>{6}</td>" + "<td>{7}</td>" + "</tr>", fullItem["Case ID"],
                fullItem["Customer Company Name"], fullItem["Customer Name"],
                fullItem["Customer Contact Email"], fullItem["Premier/BC"], 
                fullItem["TAM Name"], fullItem["Calling Country"], fullItem["Contract Country"]);
            }
            sb.Append("</table>");

            // Send the email
            SendEmail(to,subject,sb.ToString());
        }
        public static void SendEmail(string to, string subject, string body)
        {
            SendEmail(new string[] { to }, new string[] { }, subject, body);
        }
        public static void SendEmail(string[] to, string[] cc, string subject, string body)
        {
            if (to == null || to.Length == 0)
            {
                return;
            }

            body = "Hi,<br/><br/>" +
                body +
                "<br/>Azure Outage Portal System V2<br/><br/>" +
                "Note: this is an unmonitored email alias. For any feedback or questions, please email us at <a href=\"mailto:raasstatustool@microsoft.com\">***@microsoft.com</a>";

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtphost");

            //mail.From = new MailAddress("yanqwa@microsoft.com");
            mail.From = new MailAddress("yanqwa@microsoft.com");

            if (to != null && to.Length > 0)
            {
                for (int i = 0; i < to.Length; i++)
                {
                    mail.To.Add(to[i]);
                }
            }

            if (cc != null && cc.Length > 0)
            {
                for (int i = 0; i < cc.Length; i++)
                {
                    mail.CC.Add(cc[i]);
                }
            }
            mail.Subject = subject;
            mail.Body = body;
            mail.IsBodyHtml = true;

            SmtpServer.Credentials = CredentialCache.DefaultNetworkCredentials;
            SmtpServer.EnableSsl = false;

            try
            {
                SmtpServer.Send(mail);
            }
            catch (Exception ex)
            {
              
                    string.Format("Failed to send email. To: {0}, subject: {1}, message: {2}", to, subject, ex.Message);
            }
        }


        public static void ImportInfor1()
        {
            string siteUrl = "http://azportal";
            using (SPSite site = new SPSite(siteUrl))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPList list = web.GetList("http://azportal/Lists/Azure Outage List");

                    // Get the item collections which the case details haven't been updated
                    SPQuery query = new SPQuery();
                    query.Query = "<Where><Neq><FieldRef Name=\"Active\" /><Value Type=\"Text\">" + "0" + "</Value></Neq></Where>";// active!=0
                    SPListItemCollection items = list.GetItems(query);

                    // Get info and fill in all items in the item collection
                    foreach (SPListItem item in items)
                    {
                        if ((string)item["Case ID"] == null)
                            continue;

                        string SRNumber = (item["Case ID"]).ToString().Trim();
                        string MSSolveBaseURL = @"https://mssolveweb.partners.extranet.microsoft.com/MSSolveWeb/Home";

                        try
                        {
                            // call the MSSolve Web Service to get the response.
                            WebRequest webRequest = WebRequest.Create(MSSolveBaseURL + "/GetSR/" + SRNumber + "/0");
                            webRequest.Timeout = 180000;
                            ((HttpWebRequest)webRequest).UserAgent = @"Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko";
                            webRequest.UseDefaultCredentials = false;
                            webRequest.Credentials = CredentialCache.DefaultCredentials; // Use t-zefu credential
                            string text = string.Empty;
                            WebResponse response = webRequest.GetResponse();
                            Stream responseStream = response.GetResponseStream();
                            StreamReader streamReader = new StreamReader(responseStream);
                            text = streamReader.ReadToEnd();
                            response.Dispose();
                            responseStream.Close();
                            responseStream.Dispose();

                            // Use Newtonsoft to get the target data from json
                            JObject obj = JObject.Parse(text);
                            JObject ServiceRequest = (JObject)obj["Data"]["ServiceRequestResponseData"]["ServiceRequest"];
                            JObject EmailContact = (JObject)obj["Data"]["ServiceRequestResponseData"]["Contacts"];
                            IList<JToken> emailList = EmailContact["SRContacts"].Children().ToList();

                            // Fill out the list item with the data
                            item["Case Status"] = (string)ServiceRequest["StatusName"];
                            item["Customer Company Name"] = (string)ServiceRequest["AccountIdName"];
                            item["Customer Name"] = (string)ServiceRequest["CurrentAuthorizedContactIdName"];
                            item["Customer Contact Email"] = (string)emailList[0]["PrimaryEmail"];
                            item["Premier/BC"] = (string)ServiceRequest["ServiceLevelName"];
                            item["Owner Name"] = (string)ServiceRequest["OwnerUserName"];
                            item["TAM Name"] = (string)ServiceRequest["PrimaryAccountManagerIdName"];
                            item["Calling Country"] = (string)ServiceRequest["CallingCountryCode"];
                            item["Contract Country"] = (string)ServiceRequest["ContractCountryIdName"];

                            // Get engineer working group
                            string email = (string)ServiceRequest["OwnerUserInternalEmail"];
                            if (String.IsNullOrEmpty(email))
                            {
                                item["Engineer Working Group"] = "n/a";
                                item["Owner Name"] = "Not Assigned";
                            }
                            else
                            {
                                char delimiterChars = '@';
                                string[] words = email.Split(delimiterChars);
                                string alias = words[0];
                                string queryStr = "select distinct EmployeeEmail, workgroup from vwDimEmployee where iscurrent = 'yes' and EmployeeEmail = '" + alias + "'";
                                using (SqlConnection con = new SqlConnection(@"Server=detego-ctssql;Database=ssCTSDataMart;Integrated Security=True;"))
                                {
                                    using (SqlCommand cmd = new SqlCommand(queryStr, con))
                                    {
                                        cmd.CommandTimeout = 2000;
                                        con.Open();
                                        using (SqlDataReader reader = cmd.ExecuteReader())
                                        {
                                            while (reader.Read())
                                            {
                                                string fields = reader["workgroup"].ToString();
                                                item["Engineer Working Group"] = fields;
                                            }
                                        }
                                    }
                                }
                            }
                            item["Active"] = "0";

                            if (item != null)
                            {
                                item.Update();
                            }
                        }
                        catch(Exception es)
                        {
                            //errorMessage is not used, just in case of invalid case ID
                            string errorMessage = "case Id do not exist! or something else led to failure of getting other info";
                        }
                    }
                }
            }
        }
    }
}