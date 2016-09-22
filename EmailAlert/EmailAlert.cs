using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

using System.Net;
using System.IO;
using System.Net.Mail;

namespace EmailAlert
{
    class EmailAlert:SPJobDefinition
    {
        public EmailAlert() : base() { } 
        public EmailAlert(string jobName, SPService service)
            : base(jobName, service, null, SPJobLockType.None)
        {
            this.Title = "Azure Outage Email Alert Job";
        }
        public EmailAlert(string jobName, SPWebApplication webapp)
            : base(jobName, webapp, null, SPJobLockType.ContentDatabase)
        {
            this.Title = "Azure Outage Email Alert Job";
        }
        public override void Execute(Guid targetInstanceId)
        {
            EmailLogic();
        }

        public void EmailLogic()
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

        public void EmailContent(string LSINumber, int fullCount, SPListItemCollection fullItems)
        {
            string to = "t-zefu@microsoft.com";
            string subject = "Azpotral outage alert (LSI ID= " + LSINumber + ") Massive cases in one outage.";
            string body = "Welcome to AzPortal.<br/>There are " + fullCount + " cases under one LSI: " + LSINumber +
                "<br/>For detailed LSI description, please go to http://iridias/reporting/incidentlookup/" + LSINumber +
                "<br/><br/>";
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat(body);
            // Case info table
            sb.Append("<table border=\"1\">");
            sb.AppendFormat("<tr><td>Case ID</td>" + "<td>Customer Company Name</td>" +
                "<td>Customer Name</td>" + "<td>Customer Contact Email</td>" +
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
            SendEmail(to, subject, sb.ToString());
        }

        public void SendEmail(string to, string subject, string body)
        {
            SendEmail(new string[] { to }, new string[] { }, subject, body);
        }
        public void SendEmail(string[] to, string[] cc, string subject, string body)
        {
            if (to == null || to.Length == 0)
            {
                return;
            }

            body = "Hi,<br/><br/>" +
                body +
                "<br/>Azure Outage Portal System V2<br/><br/>" +
                "Note: this is an unmonitored email alias. For any feedback or questions, please email us at <a href=\"mailto:azurestatustool@microsoft.com\">***@microsoft.com</a>";

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtphost");

            mail.From = new MailAddress("yanqwa@microsoft.com");
            //mail.From = new MailAddress("t-zefu@microsoft.com");

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


    }
}
