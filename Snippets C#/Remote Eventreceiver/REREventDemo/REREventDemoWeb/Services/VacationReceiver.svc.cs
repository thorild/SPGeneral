using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System.Security;

namespace REREventDemoWeb.Services
{
    public class VacationReceiver : IRemoteEventService
    {
        /// <summary>
        /// Handles events that occur before an action occurs, such as when a user adds or deletes a list item.
        /// </summary>
        /// <param name="properties">Holds information about the remote event.</param>
        /// <returns>Holds information returned from the remote event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            SPRemoteEventResult result = new SPRemoteEventResult();

            using (var cc = GetContext())
            {
                cc.Load(cc.Web, w => w.Title);
                cc.ExecuteQuery();
                Console.WriteLine(cc.Web.Title);

                var vacationList = cc.Web.Lists.GetByTitle("Vacation");
                var item = vacationList.GetItemById(1);
                item["Demo"] = "yyy";
                item.Update();

                cc.ExecuteQuery();


                //var listItems = vacationList.GetItems(CamlQuery.CreateAllItemsQuery());
                //cc.Load(listItems);
                //cc.ExecuteQuery();

                //foreach (var v in listItems)
                //{
                //        Console.WriteLine(v["Demo"].ToString());
                //}

            }

            return result;
        }

        /// <summary>
        /// Handles events that occur after an action occurs, such as after a user adds an item to a list or deletes an item from a list.
        /// </summary>
        /// <param name="properties">Holds information about the remote event.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            using (var cc = GetContext())
            {
                cc.Load(cc.Web, w => w.Title);
                cc.ExecuteQuery();
                Console.WriteLine(cc.Web.Title);

                var vacationList = cc.Web.Lists.GetByTitle("Vacation");
                var item = vacationList.GetItemById(1);
                item["Demo"] = "xxx " + properties.ItemEventProperties.ListItemId;
                item.Update();

                cc.ExecuteQuery();


                //var listItems = vacationList.GetItems(CamlQuery.CreateAllItemsQuery());
                //cc.Load(listItems);
                //cc.ExecuteQuery();

                //foreach (var v in listItems)
                //{
                //        Console.WriteLine(v["Demo"].ToString());
                //}

            }
        }


        static ClientContext GetContext()
        {

            string adminUrl = "https://*********.sharepoint.com/";
            string userName = "************.onmicrosoft.com";
            SecureString password = GetPassword("*********");

            using (var cc = new ClientContext(adminUrl))
            {
                cc.Credentials = new SharePointOnlineCredentials(userName, password);

                return cc;
            }

        }

        private static SecureString GetPassword(string password)
        {
            SecureString pw = new SecureString();
            foreach (char c in password.ToCharArray())
            {
                pw.AppendChar(c);
            }
            return pw;
        }
    }
}
