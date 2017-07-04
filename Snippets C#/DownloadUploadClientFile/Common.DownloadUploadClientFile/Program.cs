using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Common.DownloadUploadClientFile
{
    class Program
    {

        static string url = "https://***.sharepoint.com/sites/****";
        static string user = "fredrik.thorild@s**com";
        static string password = "***";

        static string fileToGet = "https://***.sharepoint.com/sites/***/SiteAssets/Scripts/baseScripts.js";
        static string filePathOnDisc = @"c:\temp\baseScripts.js";
        static string libraryName = "Site Assets";
        static string savePath = @"/sites/***/SiteAssets/scripts/baseScripts.js";




        static void Main(string[] args)
        {
            while (true)
            {
                Console.Clear();
                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine("Download or upload? (D/U)");
                Console.ForegroundColor = ConsoleColor.White;
                var method = Console.ReadLine();

                if (method.ToLower() == "d")
                {
                    Download();
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine("Downloaded!");
                }
                else if (method.ToLower() == "u")
                {
                    Upload();
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine("Uploaded!");
                }

                System.Threading.Thread.Sleep(7000);
            }
        }


        static void Download()
        {
            int index = fileToGet.LastIndexOf('/');
            string fileName = fileToGet.Substring(index + 1);

            WebClient client = new WebClient();
            client.Credentials = GetHumanCredentials();
            client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            client.Headers.Add("User-Agent: Other");
            client.DownloadFile(fileToGet, filePathOnDisc);
        }

        static void Upload()
        {

            byte[] b = System.IO.File.ReadAllBytes(filePathOnDisc);

            using (var cc = GetContext())
            {
                var pubList = cc.Web.GetListByTitle(libraryName);
                var blubb = pubList.RootFolder.Files.Add(new FileCreationInformation() { Content = b, Url = savePath, Overwrite = true });
                cc.Load(blubb);
                cc.ExecuteQuery();
            }
        }

        public static ClientContext GetContext()
        {

            string adminUrl = url;
            string userName = user;

            SecureString securePassword = new SecureString();
            string psw = password;
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

        public static SharePointOnlineCredentials GetHumanCredentials()
        {
            string userName = user;

            SecureString securePassword = new SecureString();
            string psw = password;
            foreach (char c in psw)
            {
                securePassword.AppendChar(c);
            }

            return new SharePointOnlineCredentials(userName, securePassword);
        }


    }
}
