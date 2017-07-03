using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ThorildSnippets
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var cc = GetContext("http://mySharePointUrl"))
            {
                Console.WriteLine(cc.Web.Title);
                CreateStartPageAndAddWebPart(cc);
            }
        }

        public static ClientContext GetContext(string url)
        {

           
            string userName = "*****.*****@myUrl.onmicrosoft.com";

            SecureString securePassword = new SecureString();
            string psw = "*********";
            foreach (char c in psw)
            {
                securePassword.AppendChar(c);
            }

            try
            {

                using (var cc = new ClientContext(url))
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
        private static void CreateStartPageAndAddWebPart(ClientContext cc)
        {
            string webHomePage = DateTime.Now.Ticks + "CustomHome.aspx";
            cc.Web.AddWikiPageByUrl("/sites/TechX/SitePages/" + webHomePage);
            cc.Web.AddLayoutToWikiPage("SitePages", WikiPageLayout.ThreeColumns, webHomePage);
            WebPartEntity scriptEditorWp = new WebPartEntity();
            scriptEditorWp.WebPartXml = CustomWebPart();
            scriptEditorWp.WebPartIndex = 1;
            cc.Web.AddWebPartToWikiPage("SitePages", scriptEditorWp, webHomePage, 1, 1, false);
            cc.Web.AddWebPartToWikiPage("SitePages", scriptEditorWp, webHomePage, 1, 2, false);

            cc.Web.AddWebPartToWikiPage("SitePages", scriptEditorWp, webHomePage, 1, 3, false);

            cc.Web.SetHomePage("SitePages/" + webHomePage);
        }
        private static void ManageFeatures(ClientContext cc)
        {
            cc.Web.DeactivateFeature(Constants.MINIMALDOWNLOADSTRATEGYFEATUREID);
        }
        private static string CustomWebPart()
        {
            StringBuilder sb = new StringBuilder(20);
            sb.Append("	<webParts>	");
            sb.Append("	  <webPart xmlns=\"http://schemas.microsoft.com/WebPart/v3\">	");
            sb.Append("	    <metaData>	");
            sb.Append("	      <type name=\"Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\" />	");
            sb.Append("	      <importErrorMessage>Det går inte att importera den här webbdelen.</importErrorMessage>	");
            sb.Append("	    </metaData>	");
            sb.Append("	    <data>	");
            sb.Append("	      <properties>	");
            sb.Append("	        <property name=\"ExportMode\" type=\"exportmode\">All</property>	");
            sb.Append("	        <property name=\"HelpUrl\" type=\"string\" />	");
            sb.Append("	        <property name=\"Hidden\" type=\"bool\">False</property>	");
            sb.Append("	        <property name=\"Description\" type=\"string\">Gör att skribenter kan lägga till HTML-kodfragment eller skript.</property>	");
            sb.Append("	        <property name=\"Content\" type=\"string\">&lt;div&gt;Hej&lt;/div&gt;	");
            sb.Append("		");
            sb.Append("	</property>	");
            sb.Append("	        <property name=\"CatalogIconImageUrl\" type=\"string\" />	");
            sb.Append("	        <property name=\"Title\" type=\"string\">Min custom webpart</property>	");
            sb.Append("	        <property name=\"AllowHide\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"AllowMinimize\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"AllowZoneChange\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"TitleUrl\" type=\"string\" />	");
            sb.Append("	        <property name=\"ChromeType\" type=\"chrometype\">TitleOnly</property>	");
            sb.Append("	        <property name=\"AllowConnect\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"Width\" type=\"unit\" />	");
            sb.Append("	        <property name=\"Height\" type=\"unit\" />	");
            sb.Append("	        <property name=\"HelpMode\" type=\"helpmode\">Navigate</property>	");
            sb.Append("	        <property name=\"AllowEdit\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"TitleIconImageUrl\" type=\"string\" />	");
            sb.Append("	        <property name=\"Direction\" type=\"direction\">NotSet</property>	");
            sb.Append("	        <property name=\"AllowClose\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"ChromeState\" type=\"chromestate\">Normal</property>	");
            sb.Append("	      </properties>	");
            sb.Append("	    </data>	");
            sb.Append("	  </webPart>	");
            sb.Append("	</webParts>	");
            return sb.ToString();
        }
        private static void SiteSearch(ClientContext cc)
        {
            var searchResult = cc.Web.SiteSearch("*");
            foreach (var result in searchResult)
            {
                Console.WriteLine(result.Title);
            }
        }
        private static void InjectJavascript(ClientContext cc)
        {
            var scriptCustomAction = cc.Site.UserCustomActions.Add();
            scriptCustomAction.Location = "ScriptLink";
            scriptCustomAction.Sequence = 11;
            scriptCustomAction.ScriptSrc = "https://acmebiz.sharepoint.com/sites/TechX/SiteAssets/braJavacript.js";
            scriptCustomAction.Update();
            cc.ExecuteQuery();

        }
        private static void InjectCSS(ClientContext cc)
        {
            var scriptCustomAction = cc.Site.UserCustomActions.Add();
            scriptCustomAction.Location = "ScriptLink";
            scriptCustomAction.Sequence = 100;
            var css = "https://xxxxxx.sharepoint.com/sites/SEPSDCDevTopSite/SiteAssets/Css/";
            scriptCustomAction.ScriptBlock = @"document.write('<link rel=""stylesheet"" href=""" + css + @"/breadcrumb.css"" />');";
            scriptCustomAction.Update();
            cc.ExecuteQuery();

        }
        private static void EjectJavaScript(ClientContext cc)
        {
            var existingActions = cc.Web.UserCustomActions;
            cc.Load(existingActions);
            cc.ExecuteQuery();
            var actions = existingActions.ToArray();

            foreach (var action in actions)
            {
                if (action.Location == "ScriptLink" && action.ScriptSrc == "https://xxxxxxx.sharepoint.com/sites/vdemo/SiteAssets/myscript.js")
                {
                    action.DeleteObject();
                    cc.ExecuteQuery();
                }
            }
        }
        private static void EjectCSS(ClientContext cc)
        {
            var existingActions = cc.Web.UserCustomActions;
            cc.Load(existingActions);
            cc.ExecuteQuery();
            var actions = existingActions.ToArray();

            foreach (var action in actions)
            {
                if (action.Location == "ScriptLink" && action.ScriptBlock.Contains("stajl.css"))
                {
                    action.DeleteObject();
                    cc.ExecuteQuery();
                }
            }
        }
    }
}
