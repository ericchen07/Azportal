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

namespace OutageActiveRCAUpdate
{
    class OutageActiveRCAUpdate : SPJobDefinition
    {
        public OutageActiveRCAUpdate() : base() { }
        public OutageActiveRCAUpdate(string jobName, SPService service)
            : base(jobName, service, null, SPJobLockType.None)
        {
            this.Title = "LSI Active RCA Update Job";
        }
        public OutageActiveRCAUpdate(string jobName, SPWebApplication webapp)
            : base(jobName, webapp, null, SPJobLockType.ContentDatabase)
        {
            this.Title = "LSI Active RCA Update Job";
        }
        public override void Execute(Guid targetInstanceId)
        {
            ActiveRCAUpdateLogic();
        }

        public void ActiveRCAUpdateLogic()
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
                            caseitem["LSI Active"] = item["LSI Active"];
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
    }
}
