using Microsoft.UI.Xaml.Media;
using System;

namespace SyncOutlookCalendar
{
    public class Meeting
    {
        public string EventName { get; set; }
        public DateTime From { get; set; }
        public DateTime To { get; set; }
        public bool IsAllDay { get; set; }
        public Brush Background { get; set; }
        public Brush Foreground { get; set; }
        public string RRule { get; set; }
    }
}
