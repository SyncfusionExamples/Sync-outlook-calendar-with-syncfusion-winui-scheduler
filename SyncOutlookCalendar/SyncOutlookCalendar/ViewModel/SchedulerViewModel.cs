using System.Collections.ObjectModel;
using System;
using System.Windows.Input;
using Syncfusion.UI.Xaml.Core;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Linq;
using System.Net.Http.Headers;
using Microsoft.UI.Xaml.Media;
using Windows.UI;
using Syncfusion.UI.Xaml.Scheduler;
using MicrosoftDayOfWeek = Microsoft.Graph.DayOfWeek;
using SchedulerRecurrenceRange = Syncfusion.UI.Xaml.Scheduler.RecurrenceRange;
using SchedulerRecurrenceType = Syncfusion.UI.Xaml.Scheduler.RecurrenceType;
using System.Collections.Generic;

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

        /// <summary>
        /// Method to add events to scheduler.
        /// </summary>
        private void AddSchedulerEvents()
        {
            var colors = new List<SolidColorBrush>
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

            var subjects = new List<string>
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
                meeting.EventName = subjects[ran.Next(0, subjects.Count)];
                meeting.From = DateTime.Now.Date.AddDays(startdate).AddHours(9);
                meeting.To = meeting.From.AddHours(10);
                meeting.Background = colors[ran.Next(0, colors.Count)];
                meeting.Foreground = GetAppointmentForeground(meeting.Background);
                this.Meetings.Add(meeting);
            }
        }

        private Brush GetAppointmentForeground(Brush backgroundColor)
        {
            var brush = backgroundColor as SolidColorBrush;

            if (brush.Color.ToString().Equals("#FF8551F2") || brush.Color.ToString().Equals("#FF5363FA") || brush.Color.ToString().Equals("#FF2D99FF"))
                return new SolidColorBrush(Microsoft.UI.Colors.White);
            else
                return new SolidColorBrush(Color.FromArgb(255, 51, 51, 51));
        }

        /// <summary>
        /// Method to import the Outlook Calendar to Syncfusion Scheduler.
        /// </summary>
        /// <param name="parameter"></param>
        public void ExecuteImportCommand(object parameter)
        {
            this.Authenticate(true);
        }

        /// <summary>
        /// Method to export the Syncfusion Scheduler events to Outlook Calendar.
        /// </summary>
        /// <param name="parameter"></param>
        public void ExecuteExportCommand(object parameter)
        {
            this.Authenticate(false);
        }

        /// <summary>
        /// Method to connect application authentication with Microsoft Azure. 
        /// </summary>
        /// <param name="import">import or export events</param>
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
        /// Method to add event to outlook calendar.
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
                //// Request to add syncfusion scheduler event to the outlook calendar events.
                Client.Me.Events.Request().AddAsync(calEvent);
            }
        }

        /// <summary>
        /// Method to get outlook calendar events.
        /// </summary>
        private void GetOutlookCalendarEvents()
        {
            //// Request to get the outlook calendar events.
            var events = Client.Me.Events.Request().GetAsync().Result.ToList();
            if (events != null && events.Count > 0)
            {
                foreach (Event appointment in events)
                {
                    Meeting meeting = new Meeting()
                    {
                        EventName = appointment.Subject,
                        From = Convert.ToDateTime(appointment.Start.DateTime),
                        To = Convert.ToDateTime(appointment.End.DateTime),
                        IsAllDay = (bool)appointment.IsAllDay,
                    };

                    if (appointment.Recurrence != null)
                    {
                        AddRecurrenceRule(appointment, meeting);
                    }
                    this.Meetings.Add(meeting);
                }
            }
        }

        /// <summary>
        /// Method to update recurrence rule to appointments.
        /// </summary>
        /// <param name="appointment"></param>
        /// <param name="meeting"></param>
        private static void AddRecurrenceRule(Event appointment, Meeting meeting)
        {
            // Creating recurrence rule
            RecurrenceProperties recurrenceProperties = new RecurrenceProperties();
            if (appointment.Recurrence.Pattern.Type == RecurrencePatternType.Daily)
            {
                recurrenceProperties.RecurrenceType = SchedulerRecurrenceType.Daily;
            }
            else if (appointment.Recurrence.Pattern.Type == RecurrencePatternType.Weekly)
            {
                recurrenceProperties.RecurrenceType = SchedulerRecurrenceType.Weekly;
                foreach (var weekDay in appointment.Recurrence.Pattern.DaysOfWeek)
                {
                    if (weekDay == MicrosoftDayOfWeek.Sunday)
                    {
                        recurrenceProperties.WeekDays = WeekDays.Sunday;
                    }
                    if (weekDay == MicrosoftDayOfWeek.Monday)
                    {
                        recurrenceProperties.WeekDays = recurrenceProperties.WeekDays | WeekDays.Monday;
                    }
                    if (weekDay == MicrosoftDayOfWeek.Tuesday)
                    {
                        recurrenceProperties.WeekDays = recurrenceProperties.WeekDays | WeekDays.Tuesday;
                    }
                    if (weekDay == MicrosoftDayOfWeek.Wednesday)
                    {
                        recurrenceProperties.WeekDays = recurrenceProperties.WeekDays | WeekDays.Wednesday;
                    }
                    if (weekDay == MicrosoftDayOfWeek.Thursday)
                    {
                        recurrenceProperties.WeekDays = recurrenceProperties.WeekDays | WeekDays.Thursday;
                    }
                    if (weekDay == MicrosoftDayOfWeek.Friday)
                    {
                        recurrenceProperties.WeekDays = recurrenceProperties.WeekDays | WeekDays.Friday;
                    }
                    if (weekDay == MicrosoftDayOfWeek.Saturday)
                    {
                        recurrenceProperties.WeekDays = recurrenceProperties.WeekDays | WeekDays.Saturday;
                    }
                }
            }
            recurrenceProperties.Interval = (int)appointment.Recurrence.Pattern.Interval;
            recurrenceProperties.RecurrenceRange = SchedulerRecurrenceRange.Count;
            recurrenceProperties.RecurrenceCount = 10;
            meeting.RRule = RecurrenceHelper.CreateRRule(recurrenceProperties, meeting.From, meeting.To);
        }
    }
}
