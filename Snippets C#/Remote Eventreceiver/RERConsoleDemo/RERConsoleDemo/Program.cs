using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace RERConsoleDemo
{
    class Program
    {
        static void Main(string[] args)
        {

            //TriggerDeptUpdate();
            ConnectDisconnectRER();
        }

        static void ConnectDisconnectRER()
        {
            using (var cc = GetContext())
            {
                cc.Load(cc.Web, w => w.Title);
                cc.ExecuteQuery();
                Console.WriteLine(cc.Web.Title);

                //var vacationList = cc.Web.Lists.GetByTitle("Vacation");
                //var item = vacationList.GetItemById(1);
                //item["Demo"] = "zzz";
                //item.Update();

                //cc.ExecuteQuery();


                //var listItems = vacationList.GetItems(CamlQuery.CreateAllItemsQuery());
                //cc.Load(listItems);
                //cc.ExecuteQuery();

                //foreach (var v in listItems)
                //{
                //    Console.WriteLine(v["Title"].ToString());

                //    if (v["Avdelning"] == null)
                //    {
                //        Console.WriteLine("Avd är null oh oh");

                //        var pers = (FieldUserValue)v["Author"];

                //        Console.WriteLine(GetDepartment(cc, pers.LookupId));
                //    }
                //}



                //var x = vacationList.EventReceivers;
                //cc.Load(x);
                //cc.ExecuteQuery();


                //if (x[0].ReceiverName == "VacationReceiver")
                //{
                //    x[0].DeleteObject();
                //    cc.ExecuteQuery();

                //}

                //foreach (var v in x)
                //{
                //    Console.WriteLine(v.ReceiverName);
                //    v.DeleteObject();
                //    v.Update();
                //    vacationList.Update();
                //    cc.ExecuteQuery();
                //}


                cc.Load(cc.Web, w => w.Title);
                cc.ExecuteQuery();
                Console.WriteLine(cc.Web.Title);

                var vacationList = cc.Web.Lists.GetByTitle("vFunctions");
                cc.Load(vacationList);
                cc.ExecuteQuery();


                vacationList.EventReceivers.Add(new EventReceiverDefinitionCreationInformation
                {
                    EventType = EventReceiverType.ItemUpdated,
                    ReceiverName = "VacationListFunctionLookupItemUpdated",
                    ReceiverUrl = "https://********.azurewebsites.net/services/VacationListFunctionLookup.svc",
                    SequenceNumber = 10000
                });
                vacationList.Update();

                cc.ExecuteQuery();

                //https://**********.azurewebsites.net/services/vacationreceiver.svc
            }
        }
        static void TriggerDeptUpdate()
        {
            using (var cc = GetContext())
            {
                cc.Load(cc.Web, w => w.Title);
                cc.ExecuteQuery();
                Console.WriteLine(cc.Web.Title);

                var vacationList = cc.Web.Lists.GetByTitle("vStaff");
                cc.Load(vacationList);
                cc.ExecuteQuery();

                var listItems = vacationList.GetItems(CamlQuery.CreateAllItemsQuery());
                cc.Load(listItems);
                cc.ExecuteQuery();

                foreach (var v in listItems)
                {
                    v["Title"] = DateTime.Now.Ticks.ToString();
                    v.Update();
                }

                cc.ExecuteQuery();

            }
        }

        static string GetDepartment(ClientContext clientContext, int id)
        {

            Microsoft.SharePoint.Client.List spList = clientContext.Web.Lists.GetByTitle("People");
            clientContext.Load(spList);
            clientContext.ExecuteQuery();

            if (spList != null && spList.ItemCount > 0)
            {
                Microsoft.SharePoint.Client.CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml =
                   @"<View>  
                        <Query> 
                        <Where><Eq><FieldRef Name='Employee' LookupId='True' /><Value Type='Integer'>23</Value></Eq></Where> 
                        </Query>
                    </View>";

                ListItemCollection listItems = spList.GetItems(camlQuery);
                clientContext.Load(listItems);
                clientContext.ExecuteQuery();
                var dep = (FieldLookupValue)listItems[0]["Department_x0020_name"];

                return dep.LookupValue;
            }

            return "DepNotFound";

        }


        static ClientContext GetContext()
        {

            string adminUrl = "https://******.sharepoint.com/";
            string userName = "*************onmicrosoft.com";
            SecureString password = GetPassword("****");

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
