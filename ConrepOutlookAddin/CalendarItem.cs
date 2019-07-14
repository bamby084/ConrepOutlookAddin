using System;

namespace ConrepOutlookAddin
{
    public class CalendarItem
    {
        public int CalendarId { get; set; }
        public DateTime? StartTime { get; set; }
        public DateTime? EndTime { get; set; }
        public string Subject { get; set; }
        public string Description { get; set; }
        public string Location { get; set; }
        public bool AllDayEvent { get; set; }
        public string Attendees { get; set; }
    }
}
