using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.UI.Xaml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using static System.Formats.Asn1.AsnWriter;


// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace SyncOutlookCalendar
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        private GraphServiceClient Client;
        private static string[] scopes = { "User.Read", "Calendars.Read", "Calendars.ReadWrite" };

        public MainWindow()
        {
            this.InitializeComponent();
        }

        private async void Authenticate(bool import)
        {
            AuthenticationResult tokenRequest;
            var accounts = await App.ClientApplication.GetAccountsAsync();
            if (accounts.Count() > 0)
            {
                tokenRequest = await App.ClientApplication.AcquireTokenSilent(scopes, accounts.FirstOrDefault())
                    .ExecuteAsync();
            }
            else
            {

                tokenRequest = await App.ClientApplication.AcquireTokenInteractive(scopes).ExecuteAsync();
            }

            Client = new GraphServiceClient("https://graph.microsoft.com/v1.0/",
                                new DelegateAuthenticationProvider(async (requestMessage) =>
                                {
                                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest.AccessToken);
                                }));
            if(import)
            {
                this.GetOutlookCalendarEvents();
            }
            else
            {
                this.AddEventToOutlookCalendar();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void AddEventToOutlookCalendar()
        {
            foreach(Meeting meeting in this.Scheduler.ItemsSource)
            {
                var calEvent = new Event
                {
                    Subject = meeting.EventName,

                    Start = new DateTimeTimeZone
                    {
                        DateTime = meeting.From.ToString(),
                        TimeZone = "GMT Standard Time"
                    },
                    End = new DateTimeTimeZone()
                    {
                        DateTime = meeting.To.ToString(),
                        TimeZone = "GMT Standard Time"
                    },
                };
                Client.Me.Events.Request().AddAsync(calEvent);
            }
        }

        private void GetOutlookCalendarEvents()
        {
            var events = Client.Me.Events.Request().GetAsync().Result.ToList();
            if (events != null)
            {
                foreach (Event appointment in events)
                {
                    Meeting meeting = new Meeting()
                    {
                        EventName = appointment.Subject,
                        From = Convert.ToDateTime(appointment.Start.DateTime),
                        To = Convert.ToDateTime(appointment.End.DateTime),
                    };
                    (this.Scheduler.ItemsSource as ObservableCollection<Meeting>).Add(meeting);
                }
            }
        }

        private void ImportButtonClick(object sender, RoutedEventArgs e)
        {
            this.Authenticate(true);
        }

        private void ExportButtonClick(object sender, RoutedEventArgs e)
        {
            this.Authenticate(false);
        }
    }

}
