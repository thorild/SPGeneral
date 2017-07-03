using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace WebCreateAndSetup
{
    class Program
    {
        static string adminUrl = "https://****.sharepoint.com/sites/quotations/";
        static string userName = "****@**.com";
        static SecureString password = GetPassword("*****");

        static void Main(string[] args)
        {

            using (var clientContext = new ClientContext(adminUrl))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(userName, password);
                clientContext.Load(clientContext.Web, w => w.Webs);

                Microsoft.SharePoint.Client.List spList = clientContext.Web.Lists.GetByTitle("Quotation Log");
                clientContext.Load(spList);

                clientContext.ExecuteQuery();

                if (spList != null && spList.ItemCount > 0)
                {
                    Microsoft.SharePoint.Client.CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml =
                       @"<View>  
            <Query> 
               <Where><And><Gt><FieldRef Name='Quotation_x0020_Number' /><Value Type='Number'>16000</Value></Gt><IsNull><FieldRef Name='Quote_x0020_Link' /></IsNull></And></Where> 
            </Query> 
      </View>";

                    ListItemCollection listItems = spList.GetItems(camlQuery);
                    clientContext.Load(listItems);

                    clientContext.Load(listItems,
           itema => itema.Include(
               ttt => ttt,
               ttt => ttt["Customer"],
               ttt => ttt.Id,
               ttt => ttt["Quotation_x0020_Number"]));

                    clientContext.ExecuteQuery();
                    Console.WriteLine(listItems.Count);
                    Console.ReadLine();
                    foreach (var item in listItems)
                    {
                        string cust = "";
                        FieldLookupValue ff = (FieldLookupValue)item["Customer"];
                        if (ff.LookupValue == "*")
                            cust = "--";
                        else
                            cust = ff.LookupValue;

                        Console.WriteLine(item["Quotation_x0020_Number"] + "    " + ff.LookupValue + "    " + item.Id);
                        string url = SetupQuotSite(item["Quotation_x0020_Number"].ToString(), cust, item.Id, clientContext);
                        FixUrlForItem(url, item.Id);

                    }
                }


            }

        }

        private static void FixUrlForItem(string url, int id)
        {
            FieldUrlValue furl = new FieldUrlValue();
            furl.Description = "Open Quot.";
            furl.Url = url;

            using (var cc2 = new ClientContext(adminUrl))
            {
                cc2.Credentials = new SharePointOnlineCredentials(userName, password);
                cc2.Load(cc2.Web, w => w.Webs);
                cc2.Load(cc2.Web, w => w.Lists);
                cc2.ExecuteQuery();
                var _web = cc2.Web;

                var quotList2 = _web.Lists.First(x => x.Title == "Quotation Log");
                cc2.Load(quotList2);
                cc2.ExecuteQuery();

                var item2 = quotList2.GetItemById(id);

                item2["Quote_x0020_Link"] = furl;
                item2.Update();
                cc2.ExecuteQuery();
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

        static string SetupQuotSite(string title, string customer, int id, ClientContext cc)
        {

            //SetupQuotSite(quotNum, pCName, item.Id);
            Web _w = cc.Web;
            cc.Load(_w);
            cc.ExecuteQuery();


            Web newWeb = _w.CreateWeb(title, "qoutation-" + title + "-u" + DateTime.Now.Ticks, "", "STS#0", 1033, true, false);
            cc.Load(newWeb, t => t.Title, r => r.RootFolder, wp => wp.RootFolder.WelcomePage, u => u.Url, g => g.AssociatedOwnerGroup);
            cc.ExecuteQueryRetry();
            ListCreationInformation lci = new ListCreationInformation();
            lci.Description = "";
            lci.Title = title + " - " + customer + " Documents";
            lci.TemplateType = 101;
            lci.QuickLaunchOption = QuickLaunchOptions.On;
            List newLib = newWeb.Lists.Add(lci);
            cc.Load(newLib, x => x.DefaultViewUrl);
            cc.ExecuteQuery();

            CreateQuotFolders(cc, newLib, newWeb);

            NavigationNodeCollection collQuickLaunchNode = newWeb.Navigation.QuickLaunch;

            NavigationNodeCreationInformation nn0 = new NavigationNodeCreationInformation();
            nn0.Title = "Quotation Documents";
            nn0.IsExternal = true;
            nn0.Url = newLib.DefaultViewUrl;
            nn0.AsLastNode = true;
            collQuickLaunchNode.Add(nn0);

            NavigationNodeCreationInformation nn = new NavigationNodeCreationInformation();
            nn.Title = "Edit Quotation";
            nn.IsExternal = true;
            nn.Url = "/sites/quotations/Lists/Quotation Log/EditQuot.aspx?ID=" + id + "&Source=" + System.Net.WebUtility.UrlEncode(newLib.DefaultViewUrl);
            nn.AsLastNode = true;
            collQuickLaunchNode.Add(nn);

            ///
            NavigationNodeCreationInformation nn2 = new NavigationNodeCreationInformation();
            nn2.Title = "All Quotations";
            nn2.IsExternal = true;
            nn2.Url = "/sites/quotations/Lists/Quotation Log/AllItems.aspx";
            nn2.AsLastNode = true;
            collQuickLaunchNode.Add(nn2);


            //QuotSubmit&Mode=Site
            ///
            NavigationNodeCreationInformation nn3 = new NavigationNodeCreationInformation();
            nn3.Title = "Release to Site";
            nn3.IsExternal = true;
            nn3.Url = "/sites/quotations/Lists/Quotation Log/QuotSubmit.aspx?ID=" + id + "&Source=" + System.Net.WebUtility.UrlEncode(newLib.DefaultViewUrl) + "&Mode=Site";
            nn3.AsLastNode = true;
            collQuickLaunchNode.Add(nn3);

            //QuotSubmit&Mode=KAM
            ///
            NavigationNodeCreationInformation nn4 = new NavigationNodeCreationInformation();
            nn4.Title = "Release to KAM";
            nn4.IsExternal = true;
            nn4.Url = "/sites/quotations/Lists/Quotation Log/QuotSubmit.aspx?ID=" + id + "&Source=" + System.Net.WebUtility.UrlEncode(newLib.DefaultViewUrl) + "&Mode=KAM";
            nn4.AsLastNode = true;
            collQuickLaunchNode.Add(nn4);


            ///


            NavigationNodeCreationInformation nn5 = new NavigationNodeCreationInformation();
            nn5.Title = "Convert to Customer Project";
            nn5.IsExternal = true;
            nn5.Url = "https://foo.bar=" + System.Net.WebUtility.UrlEncode(newLib.DefaultViewUrl);
            nn5.AsLastNode = true;
            collQuickLaunchNode.Add(nn5);


            cc.Load(collQuickLaunchNode);
            cc.ExecuteQuery();

            //---

            NavigationNodeCollection qlNodes = newWeb.Navigation.QuickLaunch;
            cc.Load(qlNodes);
            cc.ExecuteQuery();

            qlNodes.ToList().ForEach(node => { if (node.Title == "Recent") { node.DeleteObject(); } });
            qlNodes.ToList().ForEach(node => { if (node.Title == "Documents") { node.DeleteObject(); } });
            qlNodes.ToList().ForEach(node => { if (node.Title == "Notebook") { node.DeleteObject(); } });
            qlNodes.ToList().ForEach(node => { if (node.Title == "Home") { node.DeleteObject(); } });

            cc.ExecuteQuery();
            //returnera webb
            return newLib.DefaultViewUrl;

        }
        static void CreateQuotFolders(ClientContext cc, List newLib, Web _web)
        {

            List<string> folders = new List<string>{
            "1 RFQ",
            "2 Communication",
            "3 Drawings and technical specifications",
            "4 RFQ Project (LQG 1_1)",
            "5 Purchase material and external operations",
            "6 Calculations",
            "7 Quotation (LQG 1_2)",
            "8 Customer decision (Order or No order)",
            "9 Handover to project or production (LQG 2)",
            "98 Work in Progress",
            "99 Archive"

            };

            foreach (string folder in folders)
            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

                itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                itemCreateInfo.LeafName = folder;

                Microsoft.SharePoint.Client.ListItem newItem = newLib.AddItem(itemCreateInfo);
                newItem["Title"] = folder;
                newItem.Update();
                cc.ExecuteQuery();
            }


        }
    }
}
