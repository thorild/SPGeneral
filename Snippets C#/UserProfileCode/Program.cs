using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Social;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace UserProfileCode
{
    class Program
    {
        static void Main(string[] args)
        {

            //Onprem SharePoint!!!

            System.Net.NetworkCredential ns = new System.Net.NetworkCredential(@"domain\user", "password");
            using (var cc = new ClientContext("https://sharepointenvironment.se/"))
            {

                cc.Credentials = ns;
                cc.Load(cc.Web, w => w.Title);
                cc.ExecuteQueryRetry();
                Console.WriteLine(cc.Web.Title);

                var users = cc.Web.SiteUsers;

                cc.Load(users);



                cc.ExecuteQueryRetry();

                foreach (var user in users)
                {
                    PP(cc, user.LoginName);

                }





            }
        }


        static void PP(ClientContext clientContext, string targetUser)
        {
            PeopleManager peopleManager = new PeopleManager(clientContext);
            PersonProperties personProperties = peopleManager.GetPropertiesFor(targetUser);


            // Load the request and run it on the server.
            // This example requests only the AccountName and UserProfileProperties
            // properties of the personProperties object.
            clientContext.Load(personProperties, p => p.AccountName, p => p.UserProfileProperties);
            clientContext.ExecuteQuery();

            try
            {
                foreach (var property in personProperties.UserProfileProperties)
                {
                    if (property.Key.ToString().ToLower().Contains("personalspace"))
                    {

                        Console.WriteLine(string.Format("{0}: {1}",
                            property.Key.ToString(), property.Value.ToString()));
                        Console.WriteLine(targetUser);
                        //Download(property.Value.ToString().Replace("MThumb","Lthumb"), targetUser);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception");
            }
            //Console.ReadKey(false);
        }

        static void Download(string imageUrl, string account)
        {
            if (imageUrl != "")
            {
                try
                {
                    WebRequest requestPic = WebRequest.Create(imageUrl);
                    requestPic.Credentials = new NetworkCredential("domain\\user", "password");

                    WebResponse responsePic = requestPic.GetResponse();

                    Image webImage = Image.FromStream(responsePic.GetResponseStream()); // Error

                    webImage.Save(@"c:\temp\dummyfilename" + ".jpg");
                }
                catch (Exception e)
                { }
            }
        }
    }
}
