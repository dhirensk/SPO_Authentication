Imports System.Net
Imports System.Xml
Imports Microsoft.SharePoint.Client
Imports System.Linq

Module Module1
    'https://code.msdn.microsoft.com/office/SharePoint-Online-0bdeb2ca
    ''' <summary>
    ''' 
    ''' </summary>
    <STAThread()>
    Sub Main()
        'This example extends the SPO_AuthenticateUsingCSOM sample
        'After authenticating, this example connects to the Lists web
        'service and gets the items in the Documents library, then
        'lists them in the console window

        'Adjust this string to point to your site on Office 365
        'Dim siteURL As String = "https://teradata.sharepoint.com"
        Dim siteURL As String = "https://incidentwatcher.sharepoint.com/sites/MSteams/"
        Console.WriteLine("Opening Site: " + siteURL)

        'Call the ClaimClientContext class to do claims mode authentication
        Using ClientContext As ClientContext = ClaimsClientContext.GetAuthenticatedContext(siteURL)
            If Not ClientContext Is Nothing Then
                'We have the client context object so claims-based authentication is complete
                ClientContext.Load(ClientContext.Web)
                ClientContext.ExecuteQuery()
                'Find out about the SP.Web object
                'Dim currentemail =  Globals.ThisAddIn.Application.Session.CurrentUser.Address
                'Display the name of the SharePoint site
                Console.WriteLine(ClientContext.Web.Title)
                Dim currentemail = "dhirensk@incidentwatcher.onmicrosoft.com"

                Dim Web As Web = ClientContext.Web
                Dim list As List = Web.Lists.GetById(New Guid("6A0E6B06-24D3-4095-953B-19AF0662B978"))
                'Dim list As List = Web.Lists.GetByTitle("IncidentWatcher")
                Dim query As New CamlQuery()
                query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='PersonEmail'/>" + "<Value Type='Text'>" + currentemail + "</Value></Eq></Where></Query></View>"

                Dim listitems As ListItemCollection = list.GetItems(query)

                ClientContext.Load(listitems)
                ClientContext.ExecuteQuery()
                Console.WriteLine("hello" + Str(listitems.LongCount))
                Dim listitem As ListItem
                For Each listitem In listitems
                    Console.WriteLine("{0} {1} {2} {3} {4} {5} {6} {7} {8} {9} {10} {11}",
                    listitem("Project"),
                    listitem("EscalationLevel"),
                    CType(listitem("Person"), FieldUserValue).LookupValue,
                    CType(listitem("PersonEmail"), FieldUserValue).LookupValue,
                    CType(listitem("EscalationPerson"), FieldUserValue).LookupValue,
                    CType(listitem("EscalationPersonEmail"), FieldUserValue).LookupValue,
                    listitem("Person_x0020_Phone"),
                    listitem("IncidentLevel"),
                    listitem("Escalation_x0020_Start_x0020_Tim"),
                    listitem("StartTime"),
                    listitem("Reminder_x0020_Beep_x002d_PopUp_"),
                    listitem("Reminder_x0020_Call_x0020_At"))
                Next
                'use for debugging only
                Console.ReadKey()
            End If
        End Using

    End Sub

End Module
