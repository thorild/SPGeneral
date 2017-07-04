using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Web.Configuration;

namespace Common.DownloadUploadClientFile
{
    public class Operations
    {

        static void WebDavDownload(string filePath, string un, string pw)
        {
            int index = filePath.LastIndexOf('/');
            string fileName = filePath.Substring(index + 1);


            WebClient client = new WebClient();
            client.Credentials = GetHumanCredentials(un, pw);
            client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            client.Headers.Add("User-Agent: Other");
            client.DownloadFile("https://abb.sharepoint.com" + filePath, @"c:\temp\imshr\" + fileName);
            //byte[] fjupp = client.DownloadData("https://abb.sharepoint.com" + filePath);


            //client.UploadData("https://***********************SDocumentsArchive/", fjupp);

            //WebDavUpload("", un, pw);
        }

        static void WebDavUpload(string un, string pw)
        {

            WebClient client = new WebClient();
            client.Credentials = GetHumanCredentials(un, pw);
            client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            client.Headers.Add("User-Agent: Other");

            byte[] responseArray = client.UploadFile("https://abb.sharepoint.com/sites/SEABBIMS/IMSDocumentsArchive/", "POST", @"C:\temp\seabbims ims\Dummy.docx");

        }

        static void BYPASS(string filePath, string un, string pw)
        {
            int index = filePath.LastIndexOf('/');
            string fileName = filePath.Substring(index + 1);


            byte[] b = System.IO.File.ReadAllBytes(@"c:\temp\foo.txt");

            using (var cc = GetHumanContext("https://**********/", "******", "***"))
            {
                var web = cc.Web;
                cc.Load(web, t => t.Title);
                cc.ExecuteQuery();
                Console.WriteLine(cc.Web.Title);

                var pubList = cc.Web.GetListByTitle("Site Assets");

                var blubb = pubList.RootFolder.Files.Add(new FileCreationInformation() { Content = b, Url = @"/sites/HVDCIMS/SiteAssets/scripts/basescripts.txt", Overwrite = true });
                cc.Load(blubb);
                cc.ExecuteQuery();


            }
        }



        // Please set the following connection strings in app.config for this WebJob to run:
        // AzureWebJobsDashboard and AzureWebJobsStorage
        static void MainTemp()
        {

            //BYPASS("", "", "");

            //Console.WriteLine("DONE");
            //Console.ForegroundColor = ConsoleColor.DarkYellow;
            //Console.WriteLine("Download or upload? (D/U)");
            //Console.ForegroundColor = ConsoleColor.White;
            //var method = Console.ReadLine();

            //if (method == "D")
            //{
            //    Console.WriteLine("Enter the full download path of the file");
            //    Console.ForegroundColor = ConsoleColor.White;
            //    var downloadPath = Console.ReadLine();
            //    Console.ForegroundColor = ConsoleColor.DarkYellow;
            //    Console.WriteLine("Enter your mailaddress");
            //    Console.ForegroundColor = ConsoleColor.White;
            //    var userName = Console.ReadLine();
            //    Console.ForegroundColor = ConsoleColor.DarkYellow;
            //    Console.WriteLine("Enter your pw");
            //    Console.ForegroundColor = ConsoleColor.White;
            //    var passWord = Console.ReadLine();
            //    Console.ForegroundColor = ConsoleColor.DarkYellow;
            //    Console.WriteLine("Getting your file...");

            //    WebDavTest(downloadPath, userName, passWord);
            //}
            //else if (method == "U")
            //{
            //    Console.WriteLine("Enter the upload path of the file");
            //    Console.ForegroundColor = ConsoleColor.White;
            //    var downloadPath = Console.ReadLine();
            //    Console.ForegroundColor = ConsoleColor.DarkYellow;
            //    Console.WriteLine("Enter your mailaddress");
            //    Console.ForegroundColor = ConsoleColor.White;
            //    var userName = Console.ReadLine();
            //    Console.ForegroundColor = ConsoleColor.DarkYellow;
            //    Console.WriteLine("Enter your pw");
            //    Console.ForegroundColor = ConsoleColor.White;
            //    var passWord = Console.ReadLine();
            //    Console.ForegroundColor = ConsoleColor.DarkYellow;
            //    Console.WriteLine("Uploading your file...");
            //    WebDavTest2(downloadPath, userName, passWord);
            //}

            //using (var cc = GetFullControlContext())
            using (var cc = GetHumanContext("https://***.sharepoint.com/sites/**", "*********", "*****!"))
            //using (var cc = GetHumanContext("https://****.sharepoint.com/sites/***", "*********", "****"))

            {
                var web = cc.Web;
                cc.Load(web, t => t.Title);
                cc.ExecuteQuery();
                Console.WriteLine(cc.Web.Title);


                //GetALLIDgreater 0

                List spList = cc.Web.Lists.GetByTitle("IMS Documents Publish");
                cc.Load(spList);
                cc.ExecuteQuery();

                if (spList != null && spList.ItemCount > 0)
                {
                    Microsoft.SharePoint.Client.CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery();

                    ListItemCollection listItems = spList.GetItems(camlQuery);
                    cc.Load(listItems);
                    cc.ExecuteQuery();

                    //StreamReader sr = new StreamReader(@"c:\temp\hrids.txt");
                    var lines = System.IO.File.ReadAllLines(@"c:\temp\imsIds.txt");


                    for (int i = 0; i < listItems.Count; i++)
                    //foreach (var item in listItems)
                    {

                        try
                        {
                            var item = listItems[i];
                            cc.Load(item.File, f => f.ServerRelativeUrl, n => n.Name);
                            cc.ExecuteQuery();
                            //Console.WriteLine(item.File.ServerRelativeUrl);
                            int index = item.File.ServerRelativeUrl.LastIndexOf('/');
                            string fileName = item.File.ServerRelativeUrl.Substring(index + 1);


                            var line = lines.First(x => x.ToLower().Contains(fileName.ToLower()));

                            var theId = line.Split('|')[1];


                            Console.WriteLine(i);

                            item["DocRefID"] = int.Parse(theId);
                            cc.Load(item.File);
                            cc.Load(item);
                            item.Update();
                            item.File.Publish("Published automatically");
                            cc.ExecuteQuery();

                            //Console.WriteLine(fileName + "|" + item.Id);

                            //WebDavDownload(item.File.ServerRelativeUrl, "fredrik.thorild@se.abb.com", "Jenny9999");

                            //var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(cc, item.File.ServerRelativeUrl);
                            //WriteToPublishingLibrary(fileInfo.Stream, item.Id, item.File.Name);

                        }
                        catch (Exception)
                        {

                            Console.WriteLine("Exc");
                        }

                    }







                    //var editList = cc.Web.GetListByTitle("IMS Documents");
                    //var eitem = editList.GetItemById(272);
                    //var efile = eitem.File;

                    //var econt = efile.OpenBinaryStream();

                    //cc.Load(efile);
                    //cc.ExecuteQuery();


                    //Stream s = econt.Value;


                    //var pubList = cc.Web.GetListByTitle("IMS Documents Publish");
                    ////var pitem = pubList.GetItemById(194);
                    ////var pfile = pitem.File;

                    ////cc.Load(pfile);
                    ////cc.ExecuteQuery();

                    ////FileSaveBinaryInformation fsb = new FileSaveBinaryInformation();
                    ////fsb.ContentStream = s;

                    ////pfile.SaveBinary(fsb);
                    ////cc.ExecuteQuery();

                    ////var item = pubList.AddItem(new ListItemCreationInformation());
                    ////item["Title"] = "Foo";
                    ////item.Update();
                    ////cc.ExecuteQuery();

                    //var blubb = pubList.RootFolder.Files.Add(new FileCreationInformation() { ContentStream = econt.Value, Url = @"/sites/SEABBIMS/IMSDocumentsArchive/foo.docx", Overwrite = true });
                    //cc.Load(blubb);
                    //cc.ExecuteQuery();



                }
            }
        }

        private static void WriteToPublishingLibrary(Stream content, int originId, string filename)
        {
            using (var cc = GetFullControlContext())
            {
                var pubList = cc.Web.GetListByTitle("IMS Documents Publish");

                var newItem = pubList.RootFolder.Files.Add(new FileCreationInformation() { ContentStream = content, Url = @"/sites/SEABBIMS/IMSDocumentsArchive/" + filename, Overwrite = true });
                cc.Load(newItem);
                cc.ExecuteQuery();
            }
        }

        static ClientContext GetReadContext()
        {
            TokenHelper.ClientId = WebConfigurationManager.AppSettings.Get("ClientIdRead");
            TokenHelper.ClientSecret = WebConfigurationManager.AppSettings.Get("ClientSecretRead");



            var siteUri = new Uri("https://abb.sharepoint.com/sites/Seabbims");

            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
            var accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm);
            var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUri.OriginalString, accessToken.AccessToken);
            return clientContext;
        }

        public static ClientContext GetFullControlContext()
        {
            TokenHelper.ClientId = WebConfigurationManager.AppSettings.Get("ClientIdFull");
            TokenHelper.ClientSecret = WebConfigurationManager.AppSettings.Get("ClientSecretFull");

            var siteUri = new Uri("https://***.sharepoint.com/sites/***");

            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
            var accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm);
            var clientContext = TokenHelper.GetClientContextWithAccessToken(siteUri.OriginalString, accessToken.AccessToken);
            return clientContext;

        }

        public static ClientContext GetHumanContext(string uri, string user, string pw)
        {

            string adminUrl = uri;
            string userName = user;

            SecureString securePassword = new SecureString();
            string psw = pw;
            foreach (char c in psw)
            {
                securePassword.AppendChar(c);
            }


            try
            {

                using (var cc = new ClientContext(adminUrl))
                {
                    cc.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                    cc.Load(cc.Web, w => w.Title);
                    cc.ExecuteQuery();
                    return cc;


                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public static SharePointOnlineCredentials GetHumanCredentials(string user, string pw)
        {
            string userName = user;

            SecureString securePassword = new SecureString();
            string psw = pw;
            foreach (char c in psw)
            {
                securePassword.AppendChar(c);
            }

            return new SharePointOnlineCredentials(userName, securePassword);
        }
    }
}
