using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClientAppAuthenticate
{
    class Program
    {
        static void Main(string[] args)
        {

            //Regga app i Azure AD först...
            using (var cc = new AuthenticationManager().GetAzureADNativeApplicationAuthenticatedContext(
                "https://sogvas13.sharepoint.com/sites/vdemo/",
                "1251e0d8-b64e-44bc-81b9-4c161b0016ba",
                "https://sogvas13.sharepoint.com/sites/vdemo/",
                null

                ))
            {
                cc.Web.Title = "Västerås Demokväll";
                cc.Web.Update();
                cc.ExecuteQueryRetry();
                //SiteSearch(cc);
            }
        }

        private static void SiteSearch(ClientContext cc)
        {
            var searchResult = cc.Web.SiteSearch("*");
            foreach (var result in searchResult)
            {
                Console.WriteLine(result.Title);
            }
        }
    }
}
