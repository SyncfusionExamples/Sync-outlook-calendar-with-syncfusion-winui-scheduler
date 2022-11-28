using System.Collections.ObjectModel;
using System;
using System.Windows.Input;
using Syncfusion.UI.Xaml.Core;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Linq;
using System.Net.Http.Headers;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Documents;
using Windows.UI;

namespace SyncOutlookCalendar
{
    public class SchedulerViewModel
    {
        private GraphServiceClient Client;
        private static string[] scopes = { "User.Read", "Calendars.Read", "Calendars.ReadWrite" };

        public ObservableCollection<Meeting> Meetings { get; set; }

        public ICommand ImportButtonCommand { get; set; }

        public ICommand ExportButtonCommand { get; set; }

        public SchedulerViewModel()
        {
            Meetings = new ObservableCollection<Meeting>();
            this.ImportButtonCommand = new DelegateCommand(ExecuteImportCommand);
            this.ExportButtonCommand = new DelegateCommand(ExecuteExportCommand);
            this.AddSchedulerEvents();
        }

        private void AddSchedulerEvents()
        {
            var brush = new ObservableCollection<SolidColorBrush>
            {
                new SolidColorBrush(Color.FromArgb(255, 133, 81, 242)),
                new SolidColorBrush(Color.FromArgb(255, 140, 245, 219)),
                new SolidColorBrush(Color.FromArgb(255, 83, 99, 250)),
                new SolidColorBrush(Color.FromArgb(255, 255, 222, 133)),
                new SolidColorBrush(Color.FromArgb(255, 45, 153, 255)),
                new SolidColorBrush(Color.FromArgb(255, 253, 183, 165)),
                new SolidColorBrush(Color.FromArgb(255, 198, 237, 115)),
                new SolidColorBrush(Color.FromArgb(255, 253, 185, 222)),
                new SolidColorBrush(Color.FromArgb(255, 83, 99, 250))
            };

            var subjectCollection = new ObservableCollection<string>
            {
                "Business Meeting",
                "Conference",
                "Medical check up",
                "Performance Check",
                "Consulting",
                "Project Status Discussion",
                "Client Meeting",
                "General Meeting",
                "Yoga Therapy",
                "GoToMeeting",
                "Plan Execution",
                "Project Plan"
            };

            Random ran = new Random();
            for (int startdate = -10; startdate < 10; startdate++)
            {
                var meeting = new Meeting();
                meeting.EventName = subjectCollection[ran.Next(0, subjectCollection.Count)];
                meeting.From = DateTime.Now.Date.AddDays(startdate).AddHours(9);
                meeting.To = meeting.From.AddHours(10);
                meeting.Background = brush[ran.Next(0, brush.Count)];
                this.Meetings.Add(meeting);
            }
        }

        public void ExecuteImportCommand(object parameter)
        {
            this.Authenticate(true);
        }

        public void ExecuteExportCommand(object parameter)
        {
            this.Authenticate(false);
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
            if (import)
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
            foreach (Meeting meeting in this.Meetings)
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
                    this.Meetings.Add(meeting);
                }
            }
        }

    }
}
