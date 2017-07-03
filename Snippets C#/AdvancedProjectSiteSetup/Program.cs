using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace AdvancedProjectSiteSetup
{
    class Program
    {
        private static SecureString Pw()
        {
            SecureString securePassword = new SecureString();
            string psw = "pass@word1";
            foreach (char c in psw)
            {
                securePassword.AppendChar(c);
            }
            return securePassword;
        }
        static string adminUrl = "https://dummyab.sharepoint.com/sites/insidan/projektportal/";
        static string userName = "sogeti@dummy.se";
        static SecureString password = Pw();
        static void Main(string[] args)
        {
            //Guid g = new Guid("F2D24475-0B28-4C7C-882C-BD0164B42C86");
            //FixWP(g, 5);
            //FixWP();
            //QS();
            //AddSenOfferTaskToSalesResp("4");
            //AddSenOfferTaskToSalesResp("5");
            //CreateOrderFolders(5);
            Kat1TaskBuilder();
        }
        static void Kat1TaskAdder(List<string> subtasks, string mainTask, List spList, ClientContext cc)
        {

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem newItem = spList.AddItem(itemCreateInfo);
            newItem["Title"] = mainTask;
            newItem["DueDate"] = DateTime.Now.AddDays(10);
            newItem.Update();
            cc.ExecuteQueryRetry();
           


            cc.Load(newItem, n => n.Id);
            cc.ExecuteQueryRetry();

            foreach (var v in subtasks)
            {
                try
                {
                    ListItemCreationInformation itemCreateInfo1 = new ListItemCreationInformation();
                    ListItem newItem1 = spList.AddItem(itemCreateInfo);
                    newItem1["Title"] = v.Split(';')[0];
                    newItem1["Body"] = v.Split(';')[1];
                    newItem1["ParentID"] = newItem.Id;
                    newItem1.Update();
                }
                catch (Exception)
                {

                    ListItemCreationInformation itemCreateInfo1 = new ListItemCreationInformation();
                    ListItem newItem1 = spList.AddItem(itemCreateInfo);
                    newItem1["Title"] = v;
                    newItem1["ParentID"] = newItem.Id;
                    newItem1.Update();
                }
            }
        }
        static void Kat1TaskBuilder()
        {
            List<Tuple<string, List<string>>> projTasks = new List<Tuple<string, List<string>>>();

            projTasks.Add(new Tuple<string, List<string>>("1 Innan överlämning", new List<string> {
"	1.1	KREDITBEDÖMNING;Kontroll av att kreditbedömning är utförd enligt kreditpolicy,	",
"	1.2	SKRIFTLIG ORDER;Skriftlig order har mottagits från kund.	",
"	1.3	ÖVERGRIPANDE TIDPLAN;Tidplan som visar att leveranstiden är rimlig. Bra om man kan läsa ut konstruktionsstart, beställning till lev och leverans till kund. Görs i excel.	",
"	1.4	PROJEKT INPLANERAT;Genom SOP. Om detta inte är gjort så ska projektet planeras in med Supply Chain och Operations	",
"	1.5	REGISTRERING I MONITOR;av kundorder & projekt.	",
"	1.6	PROJEKTETKALKYL UPPDATERAD;Offertkalkylen uppdaterad inför mötet  med valutakurs enligt kursen i Monitor vid orderregistrering).	",
"	1.7	ORDERCHECKLISTA KLAR;-	",

            }));

            projTasks.Add(new Tuple<string, List<string>>("2 Överlämning inkl kontraktsgenomgång", new List<string> {
"	2.1	GENOMGÅNG PROJEKTETS BUDGET;På mötet gås offertkalkylen igenom och läggs in som budget i projektet i Monitor. 	",
"	2.2	GENOMGÅNG OSÄKERHETER;Genomgång av risker och möjligheter på fliken osäkerheter i offertkalkylen. Görs tillsammans på mötet.	",
"	2.3	KONTRAKTSGENOMGÅNG - Orderchecklista;Gå igenom orderchecklistan inkl restlista på öppna frågor och fyll på om det behövs.	",
"	2.4	KONTRAKTSGENOMGÅNG - ITP;Genomgång offererad ITP. Finns ingen offererad ITP gäller dummys standard ITP.	",
"	2.5	KONTRAKTSGENOMGÅNG - Dokumentation;Genomgång av kundkrav avseende dokumentationskrav.	",
"	2.6	KONTRAKTSGENOMGÅNG - Beräkningar;Genomgång av Kundkrav avseende beräkningar.	",
"	2.7	Godkänd överlämning  - Vår referens på projekt och order i Monitor ändrat till projektledarens namn;- 	",


            }));

            projTasks.Add(new Tuple<string, List<string>>("3 Projektuppstart", new List<string> {
"	3.1	PROJEKTGRUPP UTSEDD;I de flesta fall innehåller projektgruppen: Projektledare, projektkvalitetsansvarig, ansvarig konstruktör, inköpare. Projektet bemannas i samarbete med ansvarig chef för de olika funktionerna.	",
"	3.2	ORDERERKÄNNANDE lämnat till projketledaren;Ordererkännande OK och signerat((om det behövs).	",
"	3.3	ORDERERKÄNNANDE SKICKAT;Ordererkännadet skickas  av projektledaren till kunden tillsammans med en info om att man är projektledare och kundens kontaktperson fr o m nu. 	",
"	3.4	ORDERREGISTRERING KLAR;Avser fliken Rader och att de artiklar som behövs är registrerade samt lagda på ordern.	",
"	3.5	UPPSTARTSMÖTE - GENOMGÅNG AV PRODUKT;Genomgång av övergripande teknisk lösning på mötet. Görs om möjligt av konstruktör men i annat fall projektledaren. 	",
"	3.6	UPPSTARTSMÖTE - GENOMGÅNG KUNDKRAV;Sammanställda kundkrav från överlämningen med eventuella ändringar  gås igenom. 	",
"	3.7	UPPSTARTSMÖTE - GENOMGÅNG PRELIMINÄR TIDPLAN;Projektledaren gör en preliminär, övergripande tidplan som gås igenom på mötet för att ge projektgruppen förutsättningar att kunna planera sina delar. 	",
"	3.8	UPPSTARTSMÖTE - GENOMGÅNG PROJEKTBUDGET;Projektledaren går igenom projektets budbet med projektgruppen	",



            }));


            projTasks.Add(new Tuple<string, List<string>>("4 Planering", new List<string> {
"	4.1	PLANERING INSPEKTIONER OCH DOKUMENTATION;Lägg in inspektionerna i projektkvalitets planering. Planera även sammanställning och leverans av kvalitetsdokumentation till kund. 	",
"	4.2	PROJEKTBUDGET AVS Q;Stäm av planerade resor och uppskattning av tidsåtgång i projektet mot de kostnader för kvaliatetstid och resor som finns i budgeten. Håller vi budget? 	",
"	4.3	OSÄKERHETER Q;Görs med fördel tillsammans med projektledaren eller annan kollega. Vilka risker och möjligheter har vi avseende kvalite i detta projekt.  Dokumentera skriftligt i ett dokument eller under Kommentarer/Anteckningar	",
"	4.4	INKÖPSPLAN FRAMTAGEN;Inköpsplan finns som flik i Projektplanen. Inköpsplanen fylls i med den information som finns och uppskattningar där vi inte har leveranstider eller priser klara. Glöm inte montage. 	",
"	4.5	PROJEKTBUDGET AVS INKÖP;Stäm av nya priser och uppskattningar i inköpsplanen och mot de kostnader för materialinköp som finns i budgeten. Håller vi budget? 	",
"	4.6	OSÄKERHETER I;Görs med fördel tillsammans med projektledaren eller annan kollega. Vilka risker och möjligheter har vi avseende inköp i detta projekt. Dokumentera skriftligt i ett dokument eller under Kommentarer/Anteckningar	",
"	4.7	OSÄKERHETER K;Görs med fördel tillsammans med projektledaren eller annan kollega. Vilka risker och möjligheter har vi avseende konstruktion i detta projekt.  Tänk framförallt på kostnader. Dokumentera skriftligt i ett dokument eller under Kommentarer/Anteckningar.	",
"	4.8	KONSTRUKTIONSTID;Använd mallen för beräkning av TK tid och stäm av den mot den tid som finns i projektets budget. Meddela projektledaren resultatet.	",
"	4.9	TIDPLAN;Hur vi ska säkerställa leveranstiden till kunden visas i tidplanen. Tidplanen ska visa alla leveranser inkl dokumentation.	",
"	4.1	BUDGET;Stäm av projektbudget med stöd av punkterna ovan och rapportera avvikelser till chef för operations.	",
"	4.11	PROJEKTPLAN klar enligt mall för projekt  i kategori 1;-	",



            }));

            projTasks.Add(new Tuple<string, List<string>>("5 Konstruktionsstart", new List<string> {
"	5.1	ORDERCHECKLISTA;Uppdaterad orderchecklista gås igenom	",
"	5.2	Punkter från Checklista konstruktion;-	",

            }));

            projTasks.Add(new Tuple<string, List<string>>("6 Konstruktion klar", new List<string> {
"	6.1	Inköpsunderlag klart;-	",
"	6.2	Kunddokumentation klart;-	",
"	6.3	Punkter från Checklista konstruktion;-	",


            }));

            projTasks.Add(new Tuple<string, List<string>>("7 Sammanställning och genomgång av inköpsunderlag", new List<string> {
"	7.1	INKÖPSUNDERLAG KLART;Ex: ITP, ytbehandlingsspecar, tekniska specifikationer, ritningsförteckning och ritningar packeterar efter inköps önskemål.	",
"	7.2	INKÖP SPEGLAS MOT KUNDORDER;Kontroll gjord att artiklarna i inköpsanmodan stämmer med kundorder.	",
"	7.3	INKÖPSANMODAN SKICKAD TILL INKÖP;Länk till anmodan samt anvisning vart underlag ligger skickas till inköp.	",



            }));

            projTasks.Add(new Tuple<string, List<string>>("8 Inköp klart", new List<string> {
"	8.1	INKÖPSORDER GRANSKAD;enligt rutin för gransking av inköpsorder.	",
"	8.2	INKÖPSORDER SKICKAD;till leverantör med kopia till projektledare.	",
"	8.3	BESTÄLLNING BEKRÄFTAD;Beställning bekräftad från leverantör och registrerad i Monitor. Inköpspriser justerade på kundorder. Justera även priser i inköpsplanen.	",
"	8.4	LEVERANSTID;Tidplan från leverantör mottagen och avstämd mot projektets tidplan	",
"	8.5	TECHNICAL TRANSFER;Teknisk genomgång genomförd med leverantör om detta behövs. 	",

            }));


            projTasks.Add(new Tuple<string, List<string>>("9 Inför packning och transport", new List<string> {
"	9.1	Kundens önskemål om MÄRKNING är kända och meddelade leverantörer;-	",
"	9.2	Kunden och leverantörer har PACKLISTOR enligt kundens elle dummys standard;-	",
"	9.3	CARGO SPEC finns framtagen (om detta behövs/är ett kundkrav);-	",
"	9.4	Frakten är beställd (om vi är ansvarig aför den);-	",

            }));

            projTasks.Add(new Tuple<string, List<string>>("10 Slutinspektion", new List<string> {
"	10.1	dummys slutinspektion genomförd;-	",
"	10.2	Kundens slutinspektion är genomförd;- 	",
"	10.3	LEVERANTÖR MEDDELAD att slutinspektionen/erna är godkända och packning kan starta.;-	",

            }));

            projTasks.Add(new Tuple<string, List<string>>("11 Leveranser", new List<string> {
"	PL11:2	PRODUKTER - levererade och utlevererade ur Monitor;-	",
"	PL11:3	MONTAGE - avslutat och godkänt av kund, utlevererat i monitor.;-	",
"	PL11:4	TEKNISK DOKUMENTATION - levererat och utlevererade ur Monitor;-	",
"	PL11:5	KVALITETSDOKUMENTATION - levererat och utlevererade ur Monitor;-	",

            }));

            projTasks.Add(new Tuple<string, List<string>>("12 Fakturering", new List<string> {
"	12.1	OK ATT FAKTURERA  - Klartecken från till ekonomi att fakturera. Pre-invoice checklist skickas till E.;-	",


            }));

            projTasks.Add(new Tuple<string, List<string>>("13	Projektavslut", new List<string> {
"	13.1	PROJEKTET SLUTLEVERERAT;Bekräftelse på slutförd leverans från kund	",
"	13.2	ADMINISTRATIVT AVSLUTAT PROJEKT;Aktiviteter och projektet avslutat i Monitor. Alla kundorder slutlevererade.	",
"	13.3	EKONOMISKT RESULTAT;Slutlig uppföljning av projektets budget med utfall i stället för prognos.	",
"	13.4	KUNDNÖJDHET - Intervjuat kunden enligt gällande mall.;-	",
"	13.5	ERFARANHETSÅTERFÖRING;Erfarenhetsåterföring med projektgrupp/säljare	",
"	13.6	PROJEKTRAPPORT;Rapporterat på styrgruppsmöte	",
"	13.7	EKONOMISKT AVSLUT;Fakturor betalda av kund	",


            }));

            Kat1TaskRunner(projTasks);

        }
        private static void Kat1TaskRunner(List<Tuple<string, List<string>>> projTasks)
        {
            using (var cc = new ClientContext("https://dummyab.sharepoint.com/sites/projekt/Projekt-5/"))
            {
                cc.Credentials = new SharePointOnlineCredentials(userName, password);
                cc.Load(cc.Web, w => w.Title);
                cc.ExecuteQuery();
                Microsoft.SharePoint.Client.List spList = cc.Web.Lists.GetByTitle("Uppgifter");
                cc.Load(spList);
                cc.ExecuteQueryRetry();

                foreach (var item in projTasks)
                {
                    Kat1TaskAdder(item.Item2, item.Item1, spList,cc);
                    cc.ExecuteQueryRetry();
                }
            }
        }

        private static void CreateTaskForSalesResp(int lookupId)
        {
            throw new NotImplementedException();
        }

        static void QS()
        {

            using (var cc = new ClientContext("https://dummyab.sharepoint.com/sites/projekt/projekt-4"))
            {
                cc.Credentials = new SharePointOnlineCredentials(userName, password);

                ListCreationInformation lci = new ListCreationInformation();
                lci.Description = "";
                Guid g = new Guid("192efa95-e50c-475e-87ab-361cede5dd7f");
                lci.TemplateFeatureId = g;
                lci.Title = "QS6";
                lci.TemplateType = 170;
                List newLib = cc.Web.Lists.Add(lci);
                cc.Load(newLib, gu => gu.Id);
                cc.ExecuteQueryRetry();
                CreateQSItems(cc, newLib);
                FixWP(newLib.Id, 4);

            }
        }

        private static void CreateQSItems(ClientContext cc, List newLib)
        {

            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = newLib.AddItem(itemCreateInfo);
                oListItem["Title"] = "Ändra ProjektdataS";

                FieldUrlValue _url1 = new FieldUrlValue();
                _url1.Url = "/sites/insidan/SiteAssets/Images/snabbstartCog2.png";
                _url1.Description = "/sites/insidan/SiteAssets/Images/snabbstartCog2.png";
                oListItem["BackgroundImageLocation"] = _url1;


                FieldUrlValue _url2 = new FieldUrlValue();
                _url2.Url = "/sites/insidan/projektportal/Lists/Projekt/EditForm.aspx?ID=4";
                _url2.Description = "/sites/insidan/projektportal/Lists/Projekt/EditForm.aspx?ID=4";
                oListItem["LinkLocation"] = _url2;


                oListItem["LaunchBehavior"] = "Dialogruta";
                oListItem["TileOrder"] = 1;
                oListItem.Update();
            }

            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = newLib.AddItem(itemCreateInfo);
                oListItem["Title"] = "ProjektuppgifterS";

                FieldUrlValue _url1 = new FieldUrlValue();
                _url1.Url = "/sites/insidan/SiteAssets/Images/snabbstartTask.png";
                _url1.Description = "/sites/insidan/SiteAssets/Images/snabbstartTask.png";
                oListItem["BackgroundImageLocation"] = _url1;


                FieldUrlValue _url2 = new FieldUrlValue();
                _url2.Url = "/sites/projekt/Projekt-4/Lists/Uppgifter/AllItems.aspx";
                _url2.Description = "/sites/projekt/Projekt-4/Lists/Uppgifter/AllItems.aspx";
                oListItem["LinkLocation"] = _url2;


                oListItem["LaunchBehavior"] = "Navigering på sidan";
                oListItem["TileOrder"] = 2;
                oListItem.Update();
            }


            {
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = newLib.AddItem(itemCreateInfo);
                oListItem["Title"] = "ProjektdokumentS";

                FieldUrlValue _url1 = new FieldUrlValue();
                _url1.Url = "/sites/insidan/SiteAssets/Images/snabbstartFile.png";
                _url1.Description = "//sites/insidan/SiteAssets/Images/snabbstartFile.png";
                oListItem["BackgroundImageLocation"] = _url1;


                FieldUrlValue _url2 = new FieldUrlValue();
                _url2.Url = "/sites/projekt/Projekt-4/Projektdokument/Forms/AllItems.aspx";
                _url2.Description = "/sites/projekt/Projekt-4/Projektdokument/Forms/AllItems.aspx";
                oListItem["LinkLocation"] = _url2;


                oListItem["LaunchBehavior"] = "Navigering på sidan";
                oListItem["TileOrder"] = 3;
                oListItem.Update();
            }



            cc.ExecuteQuery();
        }

        static void FixWP(Guid QSListId, int siteId)
        {
            using (var cc = new ClientContext("https://dummyab.sharepoint.com/sites/projekt/projekt-" + siteId))
            {

                cc.Credentials = new SharePointOnlineCredentials(userName, password);
                //cc.Web.DeactivateFeature(new Guid("87294c72-f260-42f3-a41b-981a2ffce37a"));
                //cc.Web.ActivateFeature(new Guid("00bfea71-d8fe-4fec-8dad-01c19a6e4053"));

                string projectHomePage = "ProjektStart.aspx";

                cc.Web.AddWikiPageByUrl("/sites/projekt/Projekt-" + siteId + "/SitePages/" + projectHomePage);
                cc.Web.AddLayoutToWikiPage("SitePages", WikiPageLayout.OneColumn, projectHomePage);

                WebPartEntity scriptEditorWp = new WebPartEntity();
                scriptEditorWp.WebPartXml = ScriptWP(siteId);
                scriptEditorWp.WebPartIndex = 1;
                scriptEditorWp.WebPartTitle = "ScriptEditor";
                cc.Web.AddWebPartToWikiPage("SitePages", scriptEditorWp, projectHomePage, 1, 1, false);

                WebPartEntity qswp = new WebPartEntity();
                qswp.WebPartXml = ProjectShortCuts(QSListId.ToString(), siteId);
                qswp.WebPartIndex = 1;
                qswp.WebPartTitle = "Projektgenvägar";
                cc.Web.AddWebPartToWikiPage("SitePages", qswp, projectHomePage, 1, 1, false);


                cc.Web.SetHomePage("SitePages/" + projectHomePage);
            }
        }

        static string ScriptWP(int siteId)
        {
            StringBuilder sb = new StringBuilder(20);
            sb.Append("	<webParts>	");
            sb.Append("	  <webPart xmlns=\"http://schemas.microsoft.com/WebPart/v3\">												");
            sb.Append("	    <metaData>												");
            sb.Append("	  <type name=\"Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\" />												");
            sb.Append("	  <importErrorMessage>Det går inte att importera den här webbdelen.</importErrorMessage>												");
            sb.Append("	    </metaData>												");
            sb.Append("	    <data>												");
            sb.Append("	  <properties>												");
            sb.Append("	    <property name=\"ExportMode\" type=\"exportmode\">All</property>												");
            sb.Append("	    <property name=\"HelpUrl\" type=\"string\" />												");
            sb.Append("	    <property name=\"Hidden\" type=\"bool\">False</property>												");
            sb.Append("	    <property name=\"Description\" type=\"string\">Gör att skribenter kan lägga till HTML-kodfragment eller skript.</property>												");
            sb.Append("	    <property name=\"Content\" type=\"string\">												");
            sb.Append("	  &lt;script src=\"https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js\"&gt;&lt;/script&gt;												");
            sb.Append("	  &lt;script src=\"https://cdnjs.cloudflare.com/ajax/libs/jquery.SPServices/2014.02/jquery.SPServices.min.js\"&gt;&lt;/script&gt;												");
            sb.Append("	  &lt;script src=\"https://dummyab.sharepoint.com/sites/projekt/SiteAssets/scripts/dummyproj.js\"&gt;&lt;/script&gt;												");
            sb.Append("													");
            sb.Append("													");
            sb.Append("	  &lt;svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:ev=\"http://www.w3.org/2001/xml-events\"												");
            sb.Append("	  xmlns:v=\"http://schemas.microsoft.com/visio/2003/SVGExtensions/\" width=\"6.69894in\" height=\"2.68034in\"												");
            sb.Append("	  viewBox=\"0 0 482.324 192.985\" xml:space=\"preserve\" color-interpolation-filters=\"sRGB\" class=\"st7\"&gt;												");
            sb.Append("	  &lt;v:documentProperties v:langID=\"1033\" v:metric=\"true\" v:viewMarkup=\"false\"&gt;												");
            sb.Append("	  &lt;v:userDefs&gt;												");
            sb.Append("	  &lt;v:ud v:nameU=\"msvNoAutoConnect\" v:val=\"VT0(1):26\"/&gt;												");
            sb.Append("	  &lt;/v:userDefs&gt;												");
            sb.Append("	  &lt;/v:documentProperties&gt;												");
            sb.Append("	  &lt;style type=\"text/css\"&gt;												");
            sb.Append("	  &lt;![CDATA[									.st1 {fill:#31859b;stroke:none;stroke-linecap:round;stroke-linejoin:round;stroke-width:0.24;cursor:pointer;cursor:hand}			");
            sb.Append("	  .st2 {fill:#ffffff;font-family:Calibri;font-size:0.833336em}												");
            sb.Append("	  .st3 {fill:#000000;font-family:Calibri;font-size:0.833336em}												");
            sb.Append("	  .st4 {stroke:#000000;stroke-linecap:round;stroke-linejoin:round;stroke-width:0.75}												");
            sb.Append("	  .st5 {fill:#92d050;stroke:none;stroke-linecap:round;stroke-linejoin:round;stroke-width:0.24;cursor:pointer;cursor:hand}												");
            sb.Append("	  .st6 {fill:none;stroke:none;stroke-linecap:round;stroke-linejoin:round;stroke-width:0.75}												");
            sb.Append("	  .st7 {fill:none;fill-rule:evenodd;font-size:12px;overflow:visible;stroke-linecap:square;stroke-miterlimit:3}												");
            sb.Append("	  ]]&gt;												");
            sb.Append("	  												");
            sb.Append("	  &lt;/style&gt;												");
            sb.Append("	  &lt;g v:mID=\"0\" v:index=\"1\" v:groupContext=\"foregroundPage\"&gt;												");
            sb.Append("	  &lt;title&gt;Projekt&lt;/title&gt;												");
            sb.Append("	  &lt;v:pageProperties v:drawingScale=\"0.0393701\" v:pageScale=\"0.0393701\" v:drawingUnits=\"24\" v:shadowOffsetX=\"8.50394\"	v:shadowOffsetY=\"-8.50394\"/&gt;											");
            sb.Append("	  &lt;v:layer v:name=\"Connector\" v:index=\"0\"/&gt;												");
            sb.Append("	  &lt;g id=\"shape2-1\" v:mID=\"2\" v:groupContext=\"shape\" transform=\"translate(18.12,-89.8254)\"&gt;												");
            sb.Append("	  &lt;title&gt;Förfrågan&lt;/title&gt;												");
            sb.Append("	  &lt;desc&gt;Förfrågan&lt;/desc&gt;												");
            sb.Append("	  &lt;v:userDefs&gt;												");
            sb.Append("	  &lt;v:ud v:nameU=\"visVersion\" v:prompt=\"\" v:val=\"VT0(15):26\"/&gt;												");
            sb.Append("	  &lt;/v:userDefs&gt;												");
            sb.Append("	  &lt;v:textBlock v:margins=\"rect(4,4,4,4)\" v:tabSpace=\"42.5197\"/&gt;												");
            sb.Append("	  &lt;v:textRect cx=\"63.7795\" cy=\"150.465\" width=\"127.56\" height=\"85.0394\"/&gt;												");
            sb.Append("	  &lt;path id=\"svgFF\" onclick=\"updateListItem(" + siteId + ",'Förfrågan');\" d=\"M0 192.98 L106.3 192.98 L127.56 150.47 L106.3 107.95 L0 107.95 L21.26 150.47 L0 192.98 Z\" class=\"st1\"/&gt;												");
            sb.Append("	  &lt;text x=\"44.06\" y=\"153.47\" class=\"st2\" v:langID=\"1053\"&gt;&lt;v:paragraph v:horizAlign=\"1\"/&gt;&lt;v:tabList/&gt;Förfrågan&lt;/text&gt;												");
            sb.Append("	  &lt;/g&gt;												");
            sb.Append("	  &lt;g id=\"shape3-4\" v:mID=\"3\" v:groupContext=\"shape\" transform=\"translate(131.506,-89.8254)\"&gt;												");
            sb.Append("	  &lt;title&gt;Offert Skickad&lt;/title&gt;												");
            sb.Append("	  &lt;desc&gt;Offert Skickad&lt;/desc&gt;												");
            sb.Append("	  &lt;v:userDefs&gt;												");
            sb.Append("	  &lt;v:ud v:nameU=\"visVersion\" v:prompt=\"\" v:val=\"VT0(15):26\"/&gt;												");
            sb.Append("	  &lt;/v:userDefs&gt;												");
            sb.Append("	  &lt;v:textBlock v:margins=\"rect(4,4,4,4)\" v:tabSpace=\"42.5197\"/&gt;												");
            sb.Append("	  &lt;v:textRect cx=\"63.7795\" cy=\"150.465\" width=\"127.56\" height=\"85.0394\"/&gt;												");
            sb.Append("	  &lt;path id=\"svgOS\" onclick=\"updateListItem(" + siteId + ",'Offert Skickad');\" d=\"M0 192.98 L106.3 192.98 L127.56 150.47 L106.3 107.95 L0 107.95 L21.26 150.47 L0 192.98 Z\" class=\"st1\"/&gt;												");
            sb.Append("	  &lt;text x=\"35.25\" y=\"153.47\" class=\"st2\" v:langID=\"1053\"&gt;&lt;v:paragraph v:horizAlign=\"1\"/&gt;&lt;v:tabList/&gt;Offert Skickad&lt;/text&gt;		&lt;/g&gt;										");
            sb.Append("	  &lt;g id=\"shape4-7\" v:mID=\"4\" v:groupContext=\"shape\" transform=\"translate(244.892,-89.8254)\"&gt;												");
            sb.Append("	  &lt;title&gt;Order&lt;/title&gt;												");
            sb.Append("	  &lt;desc&gt;Order&lt;/desc&gt;												");
            sb.Append("	  &lt;v:userDefs&gt;												");
            sb.Append("	  &lt;v:ud v:nameU=\"visVersion\" v:prompt=\"\" v:val=\"VT0(15):26\"/&gt;												");
            sb.Append("	  &lt;/v:userDefs&gt;												");
            sb.Append("	  &lt;v:textBlock v:margins=\"rect(4,4,4,4)\" v:tabSpace=\"42.5197\"/&gt;												");
            sb.Append("	  &lt;v:textRect cx=\"63.7795\" cy=\"150.465\" width=\"127.56\" height=\"85.0394\"/&gt;												");
            sb.Append("	  &lt;path id=\"svgOR\" onclick=\"updateListItem(" + siteId + ",'Order');\" d=\"M0 192.98 L106.3 192.98 L127.56 150.47 L106.3 107.95 L0 107.95 L21.26 150.47 L0 192.98 Z\" class=\"st1\"/&gt;												");
            sb.Append("	  &lt;text x=\"51.87\" y=\"153.47\" class=\"st2\" v:langID=\"1053\"&gt;&lt;v:paragraph v:horizAlign=\"1\"/&gt;&lt;v:tabList/&gt;Order&lt;/text&gt;												");
            sb.Append("	  &lt;/g&gt;					&lt;g id=\"shape8-10\" v:mID=\"8\" v:groupContext=\"shape\" transform=\"translate(56.9099,-38.0036)\"&gt;							");
            sb.Append("	  &lt;title&gt;Avböjd&lt;/title&gt;												");
            sb.Append("	  &lt;desc&gt;Avböjd&lt;/desc&gt;												");
            sb.Append("	  &lt;v:userDefs&gt;												");
            sb.Append("	  &lt;v:ud v:nameU=\"MaxThicknessPercent\" v:prompt=\"\" v:val=\"VT0(0.6000600060006):26\"/&gt;												");
            sb.Append("	  &lt;v:ud v:nameU=\"visVersion\" v:prompt=\"\" v:val=\"VT0(15):26\"/&gt;												");
            sb.Append("	  &lt;/v:userDefs&gt;												");
            sb.Append("	  &lt;v:textBlock v:margins=\"rect(4,4,4,4)\" v:tabSpace=\"42.5197\"/&gt;												");
            sb.Append("	  &lt;v:textRect cx=\"24.9896\" cy=\"202.987\" width=\"49.98\" height=\"20.0036\"/&gt;												");
            sb.Append("	  &lt;path id=\"svgAB\" onclick=\"updateListItem(" + siteId + ",'Avböjd');\" d=\"M49.98 168 A24.9896 24.9896 -180 1 0 0 168 A24.9896 24.9896 -180 1 0 49.98 168 ZM11.82 162.62 A13.8974 13.8974 -180 0 0 30.36 181.17 L11.82 162.62 ZM38.16 173.37 A13.8974 13.8974 -180 0 0 19.62 154.82 L38.16 173.37								 Z\" class=\"st1\"/&gt;				");
            sb.Append("	  &lt;text x=\"10.75\" y=\"205.99\" class=\"st3\" v:langID=\"1053\"&gt;&lt;v:paragraph v:horizAlign=\"1\"/&gt;&lt;v:tabList/&gt;Avböjd&lt;/text&gt;												");
            sb.Append("	  &lt;/g&gt;												");
            sb.Append("	  &lt;g id=\"shape10-13\" v:mID=\"10\" v:groupContext=\"shape\" v:layerMember=\"0\" transform=\"translate(142.136,-125.259)\"&gt;												");
            sb.Append("	  &lt;title&gt;Dynamic connector.10&lt;/title&gt;												");
            sb.Append("	  &lt;path d=\"M3.54 185.9 L10.63 185.9\" class=\"st4\"/&gt;												");
            sb.Append("	  &lt;/g&gt;												");
            sb.Append("	  &lt;g id=\"shape11-16\" v:mID=\"11\" v:groupContext=\"shape\" v:layerMember=\"0\" transform=\"translate(255.522,-125.259)\"&gt;												");
            sb.Append("	  &lt;title&gt;Dynamic connector.11&lt;/title&gt;												");
            sb.Append("	  &lt;path d=\"M3.54 185.9 L10.63 185.9\" class=\"st4\"/&gt;												");
            sb.Append("	  &lt;/g&gt;												");
            sb.Append("	  &lt;g id=\"shape15-19\" v:mID=\"15\" v:groupContext=\"shape\" transform=\"translate(421.684,-132.345)\"&gt;												");
            sb.Append("	  &lt;title&gt;Avslutad&lt;/title&gt;												");
            sb.Append("	  &lt;desc&gt;Avslutad&lt;/desc&gt;												");
            sb.Append("	  &lt;v:userDefs&gt;												");
            sb.Append("	  &lt;v:ud v:nameU=\"visVersion\" v:val=\"VT0(15):26\"/&gt;												");
            sb.Append("	  &lt;/v:userDefs&gt;												");
            sb.Append("	  &lt;v:textBlock v:margins=\"rect(4,4,4,4)\" v:tabSpace=\"42.5197\"/&gt;												");
            sb.Append("	  &lt;v:textRect cx=\"0\" cy=\"192.985\" width=\"74.41\" height=\"54\"/&gt;												");
            sb.Append("	  &lt;path id=\"svgAS\" onclick=\"updateListItem(" + siteId + ",'Avslutad');\" d=\"M-42.52 192.98 A42.5197 42.5197 0 0 1 42.52 192.98 A42.5197 42.5197 0 1 1 -42.52 192.98 Z\" class=\"st1\"/&gt;												");
            sb.Append("	  &lt;text x=\"-17.58\" y=\"195.98\" class=\"st2\" v:langID=\"1053\"&gt;&lt;v:paragraph v:horizAlign=\"1\"/&gt;&lt;v:tabList/&gt;Avslutad&lt;/text&gt;												");
            sb.Append("	  &lt;/g&gt;					&lt;g id=\"shape18-22\" v:mID=\"18\" v:groupContext=\"shape\" v:layerMember=\"0\" transform=\"translate(368.721,-125.259)\"&gt;							");
            sb.Append("	  &lt;title&gt;Dynamic connector.18&lt;/title&gt;												");
            sb.Append("	  &lt;path d=\"M3.73 185.9 L10.44 185.9\" class=\"st4\"/&gt;												");
            sb.Append("	  &lt;/g&gt;												");
            sb.Append("	  &lt;g id=\"shape20-25\" v:mID=\"20\" v:groupContext=\"shape\" transform=\"translate(244.608,-54.2004)\"&gt;												");
            sb.Append("	  &lt;title&gt;Kategori 1&lt;/title&gt;												");
            sb.Append("	  &lt;desc&gt;1&lt;/desc&gt;												");
            sb.Append("	  &lt;v:userDefs&gt;												");
            sb.Append("	  &lt;v:ud v:nameU=\"visVersion\" v:val=\"VT0(15):26\"/&gt;												");
            sb.Append("	  &lt;/v:userDefs&gt;												");
            sb.Append("	  &lt;v:textBlock v:margins=\"rect(4,4,4,4)\" v:tabSpace=\"42.5197\"/&gt;												");
            sb.Append("	  &lt;v:textRect cx=\"15.0723\" cy=\"177.913\" width=\"30.15\" height=\"30.1446\"/&gt;												");
            sb.Append("	  &lt;rect id=\"svgP1\" onclick=\"updateListItemProjModell(" + siteId + ",'Kategori 1')\" x=\"0\" y=\"162.84\" width=\"30.1446\" height=\"30.1446\" class=\"st5\"/&gt;												");
            sb.Append("	  &lt;text x=\"12.54\" y=\"180.91\" class=\"st2\" v:langID=\"1053\"&gt;&lt;v:paragraph v:horizAlign=\"1\"/&gt;&lt;v:tabList/&gt;1&lt;/text&gt;												");
            sb.Append("	  &lt;/g&gt;					&lt;g id=\"shape21-28\" v:mID=\"21\" v:groupContext=\"shape\" transform=\"translate(282.686,-54.2004)\"&gt;							");
            sb.Append("	  &lt;title&gt;Kategori 2&lt;/title&gt;								&lt;desc&gt;2&lt;/desc&gt;				");
            sb.Append("	  &lt;v:ud v:nameU=\"visVersion\" v:val=\"VT0(15):26\"/&gt;							&lt;/v:userDefs&gt;					");
            sb.Append("	  &lt;v:textBlock v:margins=\"rect(4,4,4,4)\" v:tabSpace=\"42.5197\"/&gt;												");
            sb.Append("	  &lt;v:textRect cx=\"15.0723\" cy=\"177.913\" width=\"30.15\" height=\"30.1446\"/&gt;												");
            sb.Append("	  &lt;rect id=\"svgP2\" onclick=\"updateListItemProjModell(" + siteId + ",'Kategori 2')\" x=\"0\" y=\"162.84\" width=\"30.1446\" height=\"30.1446\" class=\"st5\"/&gt;												");
            sb.Append("	  &lt;text x=\"12.54\" y=\"180.91\" class=\"st2\" v:langID=\"1053\"&gt;&lt;v:paragraph v:horizAlign=\"1\"/&gt;&lt;v:tabList/&gt;2&lt;/text&gt;												");
            sb.Append("	  &lt;/g&gt;					&lt;g id=\"shape22-31\" v:mID=\"22\" v:groupContext=\"shape\" transform=\"translate(320.763,-54.2004)\"&gt;							");
            sb.Append("	  &lt;title&gt;Kategori 3&lt;/title&gt;								&lt;desc&gt;3&lt;/desc&gt;				");
            sb.Append("	  &lt;v:ud v:nameU=\"visVersion\" v:val=\"VT0(15):26\"/&gt;							&lt;/v:userDefs&gt;					");
            sb.Append("	  &lt;v:textBlock v:margins=\"rect(4,4,4,4)\" v:tabSpace=\"42.5197\"/&gt;												");
            sb.Append("	  &lt;v:textRect cx=\"15.0723\" cy=\"177.913\" width=\"30.15\" height=\"30.1446\"/&gt;												");
            sb.Append("	  &lt;rect id=\"svgP3\" onclick=\"updateListItemProjModell(" + siteId + ",'Kategori 3')\" x=\"0\" y=\"162.84\" width=\"30.1446\" height=\"30.1446\" class=\"st5\"/&gt;												");
            sb.Append("	  &lt;text x=\"12.54\" y=\"180.91\" class=\"st2\" v:langID=\"1053\"&gt;&lt;v:paragraph v:horizAlign=\"1\"/&gt;&lt;v:tabList/&gt;3&lt;/text&gt;												");
            sb.Append("	  &lt;/g&gt;					&lt;g id=\"shape24-34\" v:mID=\"24\" v:groupContext=\"shape\" transform=\"translate(170.296,-38.0036)\"&gt;							");
            sb.Append("	  &lt;title&gt;Förlorad&lt;/title&gt;								&lt;desc&gt;Förlorad&lt;/desc&gt;				");
            sb.Append("	  &lt;v:ud v:nameU=\"MaxThicknessPercent\" v:prompt=\"\" v:val=\"VT0(0.6000600060006):26\"/&gt;												");
            sb.Append("	  &lt;v:ud v:nameU=\"visVersion\" v:prompt=\"\" v:val=\"VT0(15):26\"/&gt;							&lt;/v:userDefs&gt;					");
            sb.Append("	  &lt;v:textBlock v:margins=\"rect(4,4,4,4)\" v:tabSpace=\"42.5197\"/&gt;												");
            sb.Append("	  &lt;v:textRect cx=\"24.9896\" cy=\"202.987\" width=\"49.98\" height=\"20.0036\"/&gt;												");
            sb.Append("	  &lt;path id=\"svgFL\" onclick=\"updateListItem(" + siteId + ",'Förlorad');\" d=\"M49.98 168 A24.9896 24.9896 -180 1 0 0 168 A24.9896 24.9896 -180 1 0 49.98 168 ZM11.82 162.62 A13.8974 13.8974 -180 0 0 30.36 181.17 L11.82 162.62 ZM38.16 173.37 A13.8974 13.8974 -180 0 0 19.62 154.82 L38.16 173.37	Z\" class=\"st1\"/&gt;											");
            sb.Append("	  &lt;text x=\"7.76\" y=\"205.99\" class=\"st3\" v:langID=\"1053\"&gt;&lt;v:paragraph v:horizAlign=\"1\"/&gt;&lt;v:tabList/&gt;Förlorad&lt;/text&gt;												");
            sb.Append("	  &lt;/g&gt;					&lt;g id=\"shape25-37\" v:mID=\"25\" v:groupContext=\"shape\" transform=\"translate(244.608,-34.928)\"&gt;							");
            sb.Append("	  &lt;title&gt;Sheet.25&lt;/title&gt;								&lt;desc&gt;Projektmodell&lt;/desc&gt;				");
            sb.Append("	  &lt;v:textBlock v:margins=\"rect(4,4,4,4)\" v:tabSpace=\"42.5197\"/&gt;												");
            sb.Append("	  &lt;v:textRect cx=\"53.1496\" cy=\"185.898\" width=\"106.3\" height=\"14.1732\"/&gt;												");
            sb.Append("	  &lt;rect x=\"0\" y=\"178.812\" width=\"106.299\" height=\"14.1732\" class=\"st6\"/&gt;												");
            sb.Append("	  &lt;text x=\"24.51\" y=\"188.9\" class=\"st3\" v:langID=\"1053\"&gt;&lt;v:paragraph v:horizAlign=\"1\"/&gt;&lt;v:tabList/&gt;Projektmodell&lt;/text&gt;												");
            sb.Append("	  &lt;/g&gt;				&lt;/g&gt;							&lt;/svg&gt;	");
            sb.Append("	    </property>												");
            sb.Append("	    <property name=\"CatalogIconImageUrl\" type=\"string\" />												");
            sb.Append("	    <property name=\"Title\" type=\"string\">Skriptredigeraren</property>												");
            sb.Append("	    <property name=\"AllowHide\" type=\"bool\">True</property>												");
            sb.Append("	    <property name=\"AllowMinimize\" type=\"bool\">True</property>												");
            sb.Append("	    <property name=\"AllowZoneChange\" type=\"bool\">True</property>												");
            sb.Append("	    <property name=\"TitleUrl\" type=\"string\" />												");
            sb.Append("	    <property name=\"ChromeType\" type=\"chrometype\">None</property>												");
            sb.Append("	    <property name=\"AllowConnect\" type=\"bool\">True</property>												");
            sb.Append("	    <property name=\"Width\" type=\"unit\" />												");
            sb.Append("	    <property name=\"Height\" type=\"unit\" />												");
            sb.Append("	    <property name=\"HelpMode\" type=\"helpmode\">Navigate</property>												");
            sb.Append("	    <property name=\"AllowEdit\" type=\"bool\">True</property>												");
            sb.Append("	    <property name=\"TitleIconImageUrl\" type=\"string\" />												");
            sb.Append("	    <property name=\"Direction\" type=\"direction\">NotSet</property>												");
            sb.Append("	    <property name=\"AllowClose\" type=\"bool\">True</property>												");
            sb.Append("	    <property name=\"ChromeState\" type=\"chromestate\">Normal</property>												");
            sb.Append("	  </properties>												");
            sb.Append("	    </data>												");
            sb.Append("	  </webPart>												");
            sb.Append("	</webParts>												");

            return sb.ToString();
        }
        static string ProjectShortCuts(string listID, int siteID)
        {
            StringBuilder sb = new StringBuilder(20);
            sb.Append("	<webParts>	");
            sb.Append("	  <webPart xmlns=\"http://schemas.microsoft.com/WebPart/v3\">	");
            sb.Append("	    <metaData>	");
            sb.Append("	      <type name=\"Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\" />	");
            sb.Append("	      <importErrorMessage>Det går inte att importera den här webbdelen.</importErrorMessage>	");
            sb.Append("	    </metaData>	");
            sb.Append("	    <data>	");
            sb.Append("	      <properties>	");
            sb.Append("	        <property name=\"ShowWithSampleData\" type=\"bool\">False</property>	");
            sb.Append("	        <property name=\"Default\" type=\"string\" />	");
            sb.Append("	        <property name=\"NoDefaultStyle\" type=\"string\" />	");
            sb.Append("	        <property name=\"CacheXslStorage\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"ViewContentTypeId\" type=\"string\" />	");
            sb.Append("	        <property name=\"XmlDefinitionLink\" type=\"string\" />	");
            sb.Append("	        <property name=\"ManualRefresh\" type=\"bool\">False</property>	");
            sb.Append("	        <property name=\"ListUrl\" type=\"string\" null=\"true\" />	");
            sb.Append("	        <property name=\"ListId\" type=\"System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089\">" + listID + "</property>	");
            sb.Append("	        <property name=\"TitleUrl\" type=\"string\">/sites/projekt/Projekt-" + siteID + "/Lists/Projektgenvgar</property>	");
            sb.Append("	        <property name=\"EnableOriginalValue\" type=\"bool\">False</property>	");
            sb.Append("	        <property name=\"Direction\" type=\"direction\">NotSet</property>	");
            sb.Append("	        <property name=\"ServerRender\" type=\"bool\">False</property>	");
            sb.Append("	        <property name=\"ViewFlags\" type=\"Microsoft.SharePoint.SPViewFlags, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\">Html, TabularView, Hidden, ReadOnly, Ordered</property>	");
            sb.Append("	        <property name=\"AllowConnect\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"ListName\" type=\"string\">{" + listID + "}</property>	");
            sb.Append("	        <property name=\"ListDisplayName\" type=\"string\" null=\"true\" />	");
            sb.Append("	        <property name=\"AllowZoneChange\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"ChromeState\" type=\"chromestate\">Normal</property>	");
            sb.Append("	        <property name=\"DisableSaveAsNewViewButton\" type=\"bool\">False</property>	");
            sb.Append("	        <property name=\"ViewFlag\" type=\"string\" />	");
            sb.Append("	        <property name=\"DataSourceID\" type=\"string\" />	");
            sb.Append("	        <property name=\"ExportMode\" type=\"exportmode\">All</property>	");
            sb.Append("	        <property name=\"AutoRefresh\" type=\"bool\">False</property>	");
            sb.Append("	        <property name=\"FireInitialRow\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"AllowEdit\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"Description\" type=\"string\" />	");
            sb.Append("	        <property name=\"HelpMode\" type=\"helpmode\">Modeless</property>	");
            sb.Append("	        <property name=\"BaseXsltHashKey\" type=\"string\" null=\"true\" />	");
            sb.Append("	        <property name=\"AllowMinimize\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"CacheXslTimeOut\" type=\"int\">86400</property>	");
            sb.Append("	        <property name=\"ChromeType\" type=\"chrometype\">Default</property>	");
            sb.Append("	        <property name=\"Xsl\" type=\"string\" null=\"true\" />	");
            sb.Append("	        <property name=\"JSLink\" type=\"string\" null=\"true\" />	");
            sb.Append("	        <property name=\"CatalogIconImageUrl\" type=\"string\">/_layouts/15/images/itgen.png?rev=41</property>	");
            sb.Append("	        <property name=\"SampleData\" type=\"string\" null=\"true\" />	");
            sb.Append("	        <property name=\"UseSQLDataSourcePaging\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"TitleIconImageUrl\" type=\"string\" />	");
            sb.Append("	        <property name=\"PageSize\" type=\"int\">-1</property>	");
            sb.Append("	        <property name=\"ShowTimelineIfAvailable\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"Width\" type=\"string\" />	");
            sb.Append("	        <property name=\"DataFields\" type=\"string\" />	");
            sb.Append("	        <property name=\"Hidden\" type=\"bool\">False</property>	");
            sb.Append("	        <property name=\"Title\" type=\"string\" />	");
            sb.Append("	        <property name=\"PageType\" type=\"Microsoft.SharePoint.PAGETYPE, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\">PAGE_NORMALVIEW</property>	");
            sb.Append("	        <property name=\"DataSourcesString\" type=\"string\" />	");
            sb.Append("	        <property name=\"AllowClose\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"InplaceSearchEnabled\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"WebId\" type=\"System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089\">00000000-0000-0000-0000-000000000000</property>	");
            sb.Append("	        <property name=\"Height\" type=\"string\" />	");
            sb.Append("	        <property name=\"GhostedXslLink\" type=\"string\">main.xsl</property>	");
            sb.Append("	        <property name=\"DisableViewSelectorMenu\" type=\"bool\">False</property>	");
            sb.Append("	        <property name=\"DisplayName\" type=\"string\" />	");
            sb.Append("	        <property name=\"IsClientRender\" type=\"bool\">False</property>	");
            sb.Append("	        <property name=\"XmlDefinition\" type=\"string\">&lt;View Name=\"{0596DD30-BCA4-4BF0-A7CC-4DE93FB0FE95}\" Type=\"HTML\" Hidden=\"TRUE\" ReadOnly=\"TRUE\" OrderedView=\"TRUE\" DisplayName=\"\" Url=\"/sites/projekt/Projekt-4/SitePages/Startsida.aspx\" Level=\"1\" BaseViewID=\"1\" ContentTypeID=\"0x\" &gt;&lt;Query&gt;&lt;OrderBy&gt;&lt;FieldRef Name=\"TileOrder\" Ascending=\"TRUE\"/&gt;&lt;FieldRef Name=\"Modified\" Ascending=\"FALSE\"/&gt;&lt;/OrderBy&gt;&lt;/Query&gt;&lt;ViewFields&gt;&lt;FieldRef Name=\"Title\"/&gt;&lt;FieldRef Name=\"BackgroundImageLocation\"/&gt;&lt;FieldRef Name=\"Description\"/&gt;&lt;FieldRef Name=\"LinkLocation\"/&gt;&lt;FieldRef Name=\"LaunchBehavior\"/&gt;&lt;FieldRef Name=\"BackgroundImageClusterX\"/&gt;&lt;FieldRef Name=\"BackgroundImageClusterY\"/&gt;&lt;/ViewFields&gt;&lt;RowLimit Paged=\"TRUE\"&gt;30&lt;/RowLimit&gt;&lt;JSLink&gt;sp.ui.tileview.js&lt;/JSLink&gt;&lt;XslLink Default=\"TRUE\"&gt;main.xsl&lt;/XslLink&gt;&lt;Toolbar Type=\"Standard\"/&gt;&lt;/View&gt;</property>	");
            sb.Append("	        <property name=\"InitialAsyncDataFetch\" type=\"bool\">False</property>	");
            sb.Append("	        <property name=\"AllowHide\" type=\"bool\">True</property>	");
            sb.Append("	        <property name=\"ParameterBindings\" type=\"string\">  &lt;ParameterBinding Name=\"dvt_sortdir\" Location=\"Postback;Connection\"/&gt;	");
            sb.Append("	            &lt;ParameterBinding Name=\"dvt_sortfield\" Location=\"Postback;Connection\"/&gt;	");
            sb.Append("	            &lt;ParameterBinding Name=\"dvt_startposition\" Location=\"Postback\" DefaultValue=\"\"/&gt;	");
            sb.Append("	            &lt;ParameterBinding Name=\"dvt_firstrow\" Location=\"Postback;Connection\"/&gt;	");
            sb.Append("	            &lt;ParameterBinding Name=\"OpenMenuKeyAccessible\" Location=\"Resource(wss,OpenMenuKeyAccessible)\" /&gt;	");
            sb.Append("	            &lt;ParameterBinding Name=\"open_menu\" Location=\"Resource(wss,open_menu)\" /&gt;	");
            sb.Append("	            &lt;ParameterBinding Name=\"select_deselect_all\" Location=\"Resource(wss,select_deselect_all)\" /&gt;	");
            sb.Append("	            &lt;ParameterBinding Name=\"idPresEnabled\" Location=\"Resource(wss,idPresEnabled)\" /&gt;&lt;ParameterBinding Name=\"NoAnnouncements\" Location=\"Resource(wss,noXinviewofY_LIST)\" /&gt;&lt;ParameterBinding Name=\"NoAnnouncementsHowTo\" Location=\"Resource(wss,noXinviewofY_DEFAULT)\" /&gt;</property>	");
            sb.Append("	        <property name=\"DataSourceMode\" type=\"Microsoft.SharePoint.WebControls.SPDataSourceMode, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c\">List</property>	");
            sb.Append("	        <property name=\"AutoRefreshInterval\" type=\"int\">60</property>	");
            sb.Append("	        <property name=\"AsyncRefresh\" type=\"bool\">False</property>	");
            sb.Append("	        <property name=\"HelpUrl\" type=\"string\" />	");
            sb.Append("	        <property name=\"MissingAssembly\" type=\"string\">Det går inte att importera den här webbdelen.</property>	");
            sb.Append("	        <property name=\"XslLink\" type=\"string\" null=\"true\" />	");
            sb.Append("	        <property name=\"SelectParameters\" type=\"string\" />	");
            sb.Append("	        <property name=\"HasClientDataSource\" type=\"bool\">False</property>	");
            sb.Append("	      </properties>	");
            sb.Append("	    </data>	");
            sb.Append("	  </webPart>	");
            sb.Append("	</webParts>	");

            return sb.ToString();
        }

        static void AddContentTypesToLibrary(string leafUrl)
        {

            using (var context = new ClientContext("https://dummyab.sharepoint.com/sites/projekt/"))
            {
                context.Credentials = new SharePointOnlineCredentials(userName, password);
                Site site = context.Site;
                Web web = site.RootWeb;

                ContentType ct = web.ContentTypes.GetById("0x01010089E67208F8AB424580CC0B55F52CBA40");
                ContentType ct1 = web.ContentTypes.GetById("0x01010089E67208F8AB424580CC0B55F52CBA400102");
                ContentType ct2 = web.ContentTypes.GetById("0x01010089E67208F8AB424580CC0B55F52CBA400101");
                ContentType ct3 = web.ContentTypes.GetById("0x01010089E67208F8AB424580CC0B55F52CBA4002");
                ContentType ct4 = web.ContentTypes.GetById("0x01010089E67208F8AB424580CC0B55F52CBA4001");

                context.Load(site);
                context.Load(web);
                context.Load(ct);
                context.Load(ct1);
                context.Load(ct2);
                context.Load(ct3);
                context.Load(ct4);

                context.ExecuteQuery();

                Web subWeb = site.OpenWeb(leafUrl);

                List list = subWeb.Lists.GetByTitle("Projektdokument");
                list.ContentTypesEnabled = true;
                list.Update();
                context.ExecuteQuery();
                list.ContentTypes.AddExistingContentType(ct);
                list.ContentTypes.AddExistingContentType(ct1);
                list.ContentTypes.AddExistingContentType(ct2);
                list.ContentTypes.AddExistingContentType(ct3);
                list.ContentTypes.AddExistingContentType(ct4);

                context.Load(subWeb);
                context.Load(list);
                context.Load(ct);
                context.ExecuteQuery();

            }


            AddCustomAction(userName, password, "https://dummyab.sharepoint.com/sites/projekt/" + leafUrl + "/", "Projektdokument");

        }

        private static void AddCustomAction(string usr, SecureString psw, string url, string libName)
        {
            using (ClientContext cc = new ClientContext(url))
            {
                string title = "dummy Mallar";
                cc.Credentials = new SharePointOnlineCredentials(usr, psw);
                Web _w = cc.Web;
                cc.Load(_w, w => w.Title, l => l.Lists);
                cc.ExecuteQueryRetry();


                var spList = _w.Lists.First(x => x.Title == libName);
                cc.Load(spList);
                cc.ExecuteQueryRetry();

                UserCustomActionCollection caColl = spList.UserCustomActions;
                cc.Load(caColl);
                cc.ExecuteQuery();

                bool found = false;
                UserCustomAction newUCAToRemove = null;
                for (int i = 0; i < caColl.Count; i++)
                {
                    if (caColl[i].Title == title)
                    {
                        newUCAToRemove = caColl[i];
                        found = true;
                        break;
                    }
                }

                if (found)
                {
                    newUCAToRemove.DeleteObject();
                    Console.WriteLine("hittad");
                }

                cc.ExecuteQuery();

                UserCustomAction action = caColl.Add();
                action.Location = "CommandUI.Ribbon.ListView";
                action.Sequence = 1;
                action.Title = title;
                action.CommandUIExtension = @"<CommandUIExtension><CommandUIDefinitions>"
                       + "<CommandUIDefinition Location=\"Ribbon.Documents.New.Controls._children\">"
                       + "<Button Id=\"InvokeAction.Button\" TemplateAlias=\"o1\" Command=\"Invoke_Command\" CommandType=\"General\" LabelText=\"dummy Mallar\" Image32by32=\"https://dummyab.sharepoint.com/sites/dokumentcenter/SiteAssets/mallbild32.png\" Image16by16=\"https://dummyab.sharepoint.com/sites/dokumentcenter/SiteAssets/mallbild16.png\" />"
                       + "</CommandUIDefinition>"
                       + "</CommandUIDefinitions>"
                       + "<CommandUIHandlers>"
                       + "<CommandUIHandler Command =\"Invoke_Command\" CommandAction=\"javascript:OpenPopUpPageWithTitle('https://dummyab.sharepoint.com/sites/dokumentcenter/Mallar/Forms/AllItems.aspx?rootSaveLoc=" + url + "&amp;rootName=" + libName + "/',RefreshOnDialogClose,600,400,'dummy Mallar');\" />"
                + "</CommandUIHandlers></CommandUIExtension>";
                //OpenPopUpPageWithTitle(\"https://dummyab.sharepoint.com/sites/dokumentcenter/Mallar/Forms/AllItems.aspx?rootSaveLoc=" + url + 
                //"&rootName=" + libName + "\",RefreshOnDialogClose, 600, 400,\"dummy Mallar\")                //
                action.Update();

                cc.ExecuteQuery();

                spList.Update();
                cc.ExecuteQuery();
            }
        }

        static void CreateOrderFolders(int id)
        {
            using (var cc = new ClientContext("https://dummyab.sharepoint.com/sites/projekt/projekt-" + id))
            {
                cc.Credentials = new SharePointOnlineCredentials(userName, password);
                cc.Load(cc.Web, w => w.Title);
                cc.ExecuteQuery();
                List list = cc.Web.Lists.GetByTitle("Projektdokument");
                cc.Load(list);
                cc.ExecuteQuery();
                //    List<string> folders = new List<string>{

                //    "Beställning",
                //    "Produkt",
                //    "Projektledning",
                //    "Dokumentation"



                //};


                //    foreach (string folder in folders)
                //    {
                //        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

                //        itemCreateInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                //        itemCreateInfo.LeafName = folder;

                //        Microsoft.SharePoint.Client.ListItem newItem = list.AddItem(itemCreateInfo);
                //        newItem["Title"] = folder;
                //        newItem.Update();
                //        cc.ExecuteQueryRetry();
                //    }


                Folder fk = cc.Web.GetFolderByServerRelativeUrl("/sites/projekt/Projekt-5/Projektdokument/Kalkyl");
                Folder fr = cc.Web.GetFolderByServerRelativeUrl("/sites/projekt/Projekt-5/Projektdokument/Ritningar");
                Folder fu = cc.Web.GetFolderByServerRelativeUrl("/sites/projekt/Projekt-5/Projektdokument/Underlag från kund");
                Folder fle = cc.Web.GetFolderByServerRelativeUrl("/sites/projekt/Projekt-5/Projektdokument/Leverantörsofferter");
                Folder flj = cc.Web.GetFolderByServerRelativeUrl("/sites/projekt/Projekt-5/Projektdokument/Ljudberäkningar");

                cc.Load(fk);
                cc.Load(fr);
                cc.Load(fu);
                cc.Load(fle);
                cc.Load(flj);

                cc.ExecuteQuery();

                fk.MoveTo("/sites/projekt/Projekt-5/Projektdokument/Offert/Kalkyl");
                fr.MoveTo("/sites/projekt/Projekt-5/Projektdokument/Offert/Ritningar");
                fu.MoveTo("/sites/projekt/Projekt-5/Projektdokument/Offert/Underlag från kund");
                fle.MoveTo("/sites/projekt/Projekt-5/Projektdokument/Offert/Leverantörsofferter");
                flj.MoveTo("/sites/projekt/Projekt-5/Projektdokument/Offert/Ljudberäkningar");

                cc.ExecuteQuery();
            }
        }

        static void AddSenOfferTaskToSalesResp(string id)
        {
            using (var cc = new ClientContext("https://dummyab.sharepoint.com/sites/insidan/projektportal"))
            {
                cc.Credentials = new SharePointOnlineCredentials(userName, password);
                cc.Load(cc.Web, w => w.Title);
                cc.ExecuteQuery();

                Microsoft.SharePoint.Client.List spList = cc.Web.Lists.GetByTitle("Projekt");
                cc.Load(cc.Web, u => u.SiteUsers);
                cc.Load(spList);
                cc.ExecuteQueryRetry();

                var li = spList.GetItemById(int.Parse(id));
                cc.Load(li);
                cc.ExecuteQueryRetry();
                FieldUserValue uV = (FieldUserValue)li["S_x00e4_ljare"];

                if (uV.LookupId != null)
                    CreateTaskForSalesResp(uV.LookupId);

                Console.WriteLine(uV.LookupId + "    " + uV.LookupValue);
                string ss = uV.LookupValue;
                User user = cc.Web.EnsureUser(uV.LookupValue);
                cc.Load(user);
                cc.ExecuteQuery();

                Console.WriteLine(user.LoginName);

                var user2 = cc.Web.SiteUsers.GetById(uV.LookupId);
                cc.Load(user2);
                cc.ExecuteQuery();
                Console.WriteLine(user2.LoginName);
                //Task(ss);
            }
        }

    }
}
