using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;

namespace MS_Reusable_WP_Listing.CalendarWebPart
{
    public partial class CalendarWebPartUserControl : UserControl
    {
        #region Properties

        // Property to store the list name from the web part properties
        private string _listName = string.Empty;
        public string ListName
        {
            get { return _listName; }
            set { _listName = value; }
        }

        // Property to store the calendar title
        private string _calendarTitle = "Calendar";
        public string CalendarTitle
        {
            get { return _calendarTitle; }
            set { _calendarTitle = value; }
        }

        // Property to store the current month and year
        private DateTime _currentDate = DateTime.Today;
        public DateTime CurrentDate
        {
            get { return _currentDate; }
            set { _currentDate = value; }
        }

        // Property to store the selected date
        private DateTime _selectedDate = DateTime.Today;
        public DateTime SelectedDate
        {
            get { return _selectedDate; }
            set { _selectedDate = value; }
        }

        // Property to store the events
        private List<CalendarEvent> _events = new List<CalendarEvent>();

        #endregion

        #region Page Events

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                // Set the calendar title
                ltrCalendarTitle.Text = CalendarTitle;

                // Load events from the SharePoint list
                LoadEvents();

                // Display the calendar
                DisplayCalendar();
            }
        }

        #endregion

        #region Button Events

        protected void btnPrevMonth_Click(object sender, EventArgs e)
        {
            CurrentDate = CurrentDate.AddMonths(-1);
            DisplayCalendar();
        }

        protected void btnToday_Click(object sender, EventArgs e)
        {
            CurrentDate = DateTime.Today;
            SelectedDate = DateTime.Today;
            DisplayCalendar();
        }

        protected void btnNextMonth_Click(object sender, EventArgs e)
        {
            CurrentDate = CurrentDate.AddMonths(1);
            DisplayCalendar();
        }

        #endregion

        #region Helper Methods

        private void LoadEvents()
        {
            _events.Clear();

            try
            {
                if (string.IsNullOrEmpty(ListName))
                {
                    return;
                }

                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                using (SPWeb web = site.OpenWeb())
                {
                    // Check if the list exists
                    SPList list = web.Lists.TryGetList(ListName);
                    if (list == null)
                    {
                        return;
                    }

                    // Calculate the date range for the current month view
                    DateTime firstDayOfMonth = new DateTime(CurrentDate.Year, CurrentDate.Month, 1);
                    DateTime lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

                    // Get the first day of the calendar grid (might be in the previous month)
                    int daysToSubtract = (int)firstDayOfMonth.DayOfWeek;
                    DateTime firstDayOfCalendarGrid = firstDayOfMonth.AddDays(-daysToSubtract);

                    // Get the last day of the calendar grid (might be in the next month)
                    DateTime lastDayOfCalendarGrid = firstDayOfCalendarGrid.AddDays(42); // 6 weeks * 7 days

                    // Create a CAML query to get events in the date range
                    string camlQuery = $@"
                    <Where>
                        <And>
                            <Leq>
                                <FieldRef Name='StartDate' />
                                <Value Type='DateTime'>{SPUtility.CreateISO8601DateTimeFromSystemDateTime(lastDayOfCalendarGrid)}</Value>
                            </Leq>
                            <Geq>
                                <FieldRef Name='EndDate' />
                                <Value Type='DateTime'>{SPUtility.CreateISO8601DateTimeFromSystemDateTime(firstDayOfCalendarGrid)}</Value>
                            </Geq>
                        </And>
                    </Where>
                    <OrderBy>
                        <FieldRef Name='StartDate' Ascending='True' />
                    </OrderBy>";

                    SPQuery query = new SPQuery();
                    query.Query = camlQuery;
                    query.ViewFields = "<FieldRef Name='Title' /><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><FieldRef Name='Location' /><FieldRef Name='Description' /><FieldRef Name='Category' />";
                    query.ViewFieldsOnly = true;

                    query.Query = ApplyApprovedItemsFilter(list, query.Query);

                    SPListItemCollection items = list.GetItems(query);

                    foreach (SPListItem item in items)
                    {
                        CalendarEvent calEvent = new CalendarEvent();
                        calEvent.Id = item.ID;
                        calEvent.Title = item["Title"] != null ? item["Title"].ToString() : string.Empty;
                        
                        if (item["StartDate"] != null && DateTime.TryParse(item["StartDate"].ToString(), out DateTime startDate))
                            calEvent.StartDate = startDate;
                        
                        if (item["EndDate"] != null && DateTime.TryParse(item["EndDate"].ToString(), out DateTime endDate))
                            calEvent.EndDate = endDate;
                        
                        calEvent.Location = item["Location"] != null ? item["Location"].ToString() : string.Empty;
                        calEvent.Description = item["Description"] != null ? item["Description"].ToString() : string.Empty;
                        
                        // Handle category (used for event color)
                        if (item["Category"] != null)
                        {
                            try
                            {
                                SPFieldLookupValue lookupValue = new SPFieldLookupValue(item["Category"].ToString());
                                calEvent.Category = lookupValue.LookupValue;
                            }
                            catch
                            {
                                calEvent.Category = item["Category"].ToString();
                            }
                        }

                        _events.Add(calEvent);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error or display message
            }
        }

        private void DisplayCalendar()
        {
            // Set the current month and year
            ltrCurrentMonth.Text = CurrentDate.ToString("MMMM") + " ";
            ltrCurrentYear.Text = CurrentDate.Year.ToString();

            // Set the selected date
            ltrSelectedDate.Text = SelectedDate.ToString("MMMM d, yyyy");

            // Create the calendar days
            List<CalendarDay> calendarDays = new List<CalendarDay>();

            // Get the first day of the month
            DateTime firstDayOfMonth = new DateTime(CurrentDate.Year, CurrentDate.Month, 1);

            // Get the last day of the month
            DateTime lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

            // Get the first day of the calendar grid (might be in the previous month)
            int daysToSubtract = (int)firstDayOfMonth.DayOfWeek;
            DateTime firstDayOfCalendarGrid = firstDayOfMonth.AddDays(-daysToSubtract);

            // Create 42 days (6 weeks)
            for (int i = 0; i < 42; i++)
            {
                DateTime currentDay = firstDayOfCalendarGrid.AddDays(i);
                CalendarDay calendarDay = new CalendarDay();
                calendarDay.Date = currentDay;
                calendarDay.IsCurrentMonth = currentDay.Month == CurrentDate.Month;
                calendarDay.IsToday = currentDay.Date == DateTime.Today.Date;
                calendarDay.IsSelected = currentDay.Date == SelectedDate.Date;

                // Add events for this day
                calendarDay.Events = _events.Where(e => 
                    (e.StartDate.Date <= currentDay.Date && e.EndDate.Date >= currentDay.Date)
                ).ToList();

                calendarDays.Add(calendarDay);
            }

            // Bind the calendar days to the repeater
            rptCalendarDays.DataSource = calendarDays;
            rptCalendarDays.DataBind();

            // Bind the events for the selected date to the events repeater
            List<CalendarEvent> selectedDateEvents = _events.Where(e => 
                (e.StartDate.Date <= SelectedDate.Date && e.EndDate.Date >= SelectedDate.Date)
            ).ToList();

            rptEvents.DataSource = selectedDateEvents;
            rptEvents.DataBind();
        }

        #endregion

        #region Repeater Events

        protected void rptCalendarDays_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {
            if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
            {
                CalendarDay day = (CalendarDay)e.Item.DataItem;
                Repeater rptDayEvents = (Repeater)e.Item.FindControl("rptDayEvents");

                // Limit the number of events shown in the day cell
                int maxEventsToShow = 3;
                List<CalendarEvent> eventsToShow = day.Events.Take(maxEventsToShow).ToList();

                rptDayEvents.DataSource = eventsToShow;
                rptDayEvents.DataBind();
            }
        }

        protected void rptEvents_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {
            // Additional event binding logic can be added here if needed
        }

        private static string ApplyApprovedItemsFilter(SPList list, string existingCaml)
        {
            if (list == null || !list.EnableModeration)
                return existingCaml;

            string moderationWhere = "<Eq><FieldRef Name='_ModerationStatus' /><Value Type='ModStat'>0</Value></Eq>";

            if (string.IsNullOrEmpty(existingCaml))
                return "<Where>" + moderationWhere + "</Where>";

            string lower = existingCaml.ToLowerInvariant();
            int whereStart = lower.IndexOf("<where>", StringComparison.Ordinal);
            int whereEnd = lower.IndexOf("</where>", StringComparison.Ordinal);
            if (whereStart >= 0 && whereEnd > whereStart)
            {
                int innerStart = whereStart + "<where>".Length;
                string innerWhere = existingCaml.Substring(innerStart, whereEnd - innerStart);
                string newWhere = "<Where><And>" + innerWhere + moderationWhere + "</And></Where>";
                return existingCaml.Substring(0, whereStart) + newWhere + existingCaml.Substring(whereEnd + "</where>".Length);
            }

            return "<Where>" + moderationWhere + "</Where>" + existingCaml;
        }

        #endregion

        #region Data Binding Methods

        protected string GetDayClasses(object dataItem)
        {
            CalendarDay day = (CalendarDay)dataItem;
            string classes = "";

            if (!day.IsCurrentMonth)
                classes += " other-month";

            if (day.IsToday)
                classes += " today";

            if (day.IsSelected)
                classes += " selected";

            return classes;
        }

        protected string GetDayNumber(object dataItem)
        {
            CalendarDay day = (CalendarDay)dataItem;
            return day.Date.Day.ToString();
        }

        protected string GetEventColor(object dataItem)
        {
            CalendarEvent calEvent = (CalendarEvent)dataItem;
            
            // Map categories to colors
            Dictionary<string, string> categoryColors = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Meeting", "#4285F4" },  // Blue
                { "Holiday", "#EA4335" },  // Red
                { "Training", "#FBBC05" }, // Yellow
                { "Travel", "#34A853" },   // Green
                { "Conference", "#8E24AA" }, // Purple
                { "Personal", "#FB8C00" }  // Orange
            };

            if (!string.IsNullOrEmpty(calEvent.Category) && categoryColors.ContainsKey(calEvent.Category))
                return categoryColors[calEvent.Category];

            // Default color if no category match
            return "#777777";
        }

        protected string FormatEventTime(object dataItem)
        {
            CalendarEvent calEvent = (CalendarEvent)dataItem;
            
            // Same day event
            if (calEvent.StartDate.Date == calEvent.EndDate.Date)
            {
                return calEvent.StartDate.ToString("h:mm tt") + " - " + calEvent.EndDate.ToString("h:mm tt");
            }
            
            // Multi-day event
            return calEvent.StartDate.ToString("MMM d, h:mm tt") + " - " + calEvent.EndDate.ToString("MMM d, h:mm tt");
        }

        #endregion
    }

    #region Helper Classes

    public class CalendarDay
    {
        public DateTime Date { get; set; }
        public bool IsCurrentMonth { get; set; }
        public bool IsToday { get; set; }
        public bool IsSelected { get; set; }
        public List<CalendarEvent> Events { get; set; } = new List<CalendarEvent>();
    }

    public class CalendarEvent
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Location { get; set; }
        public string Description { get; set; }
        public string Category { get; set; }
    }

    #endregion
}
