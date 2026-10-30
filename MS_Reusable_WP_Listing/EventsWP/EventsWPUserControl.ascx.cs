using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;

namespace MS_Reusable_WP_Listing.EventsWP
{
    public partial class EventsWPUserControl : UserControl
    {
        #region Control Declarations
        
        // Manually added control declarations since they're not in the designer file
        protected global::System.Web.UI.WebControls.LinkButton btnPrevMonth;
        protected global::System.Web.UI.WebControls.LinkButton btnNextMonth;
        protected global::System.Web.UI.WebControls.Literal litCurrentMonth;
        protected global::System.Web.UI.WebControls.Literal litMonthNameDisplay;
        protected global::System.Web.UI.WebControls.Literal litYearDisplay;
        protected global::System.Web.UI.WebControls.Literal litMonthNumber;
        
        #endregion
        
        #region Properties

        private const string ViewStateKeyCurrentDate = "EventsWP_CurrentDate";
        private const string ViewStateKeySelectedDate = "EventsWP_SelectedDate";

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
        public DateTime CurrentDate
        {
            get
            {
                object value = ViewState[ViewStateKeyCurrentDate];
                return value is DateTime dateTime ? dateTime : DateTime.Today;
            }
            set
            {
                ViewState[ViewStateKeyCurrentDate] = value;
            }
        }

        // Property to store the selected date
        public DateTime SelectedDate
        {
            get
            {
                object value = ViewState[ViewStateKeySelectedDate];
                return value is DateTime dateTime ? dateTime : DateTime.Today;
            }
            set
            {
                ViewState[ViewStateKeySelectedDate] = value;
            }
        }

        // Property to store the events
        private List<CalendarEvent> _events = new List<CalendarEvent>();

        #endregion

        #region Page Events

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
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
            try
            {
                // Store current month for debugging
                int oldMonth = CurrentDate.Month;
                
                // Decrement the month
                CurrentDate = CurrentDate.AddMonths(-1);
                
                // Set selected date to the first day of the new month
                SelectedDate = new DateTime(CurrentDate.Year, CurrentDate.Month, 1);
                
                // Add debug information
                System.Web.UI.ScriptManager.RegisterStartupScript(this, this.GetType(), "PrevMonthDebug", 
                    $"console.log('Moving from month {oldMonth} to month {CurrentDate.Month} ({CurrentDate.ToString("MMMM yyyy")})');", true);
                
                // Force clear any cached data
                _events.Clear();
                    
                // Reload events for the new month
                LoadEvents();
                
                // Update the UI
                DisplayCalendar();
                
                // Force update the panel to ensure all UI elements are refreshed
                if (UpdatePanel1 != null)
                {
                    UpdatePanel1.Update();
                }
            }
            catch (Exception ex)
            {
                // Log any errors
                System.Diagnostics.Debug.WriteLine($"Error in btnPrevMonth_Click: {ex.Message}");
                System.Web.UI.ScriptManager.RegisterStartupScript(this, this.GetType(), "PrevMonthError", 
                    $"console.error('Error navigating to previous month: {ex.Message.Replace("'", "\\'")}')", true);
            }
        }

        protected void btnToday_Click(object sender, EventArgs e)
        {
            CurrentDate = DateTime.Today;
            SelectedDate = DateTime.Today;
            
            // Force clear any cached data
            _events.Clear();
            
            // Reload events for the current month
            LoadEvents();
            
            // Update the UI
            DisplayCalendar();
            
            // Force update the panel to ensure all UI elements are refreshed
            if (UpdatePanel1 != null)
            {
                UpdatePanel1.Update();
            }
        }

        protected void btnNextMonth_Click(object sender, EventArgs e)
        {
            try
            {
                // Store current month for debugging
                int oldMonth = CurrentDate.Month;
                
                // Increment the month
                CurrentDate = CurrentDate.AddMonths(1);
                
                // Set selected date to the first day of the new month
                SelectedDate = new DateTime(CurrentDate.Year, CurrentDate.Month, 1);
                
                // Add debug information
                System.Web.UI.ScriptManager.RegisterStartupScript(this, this.GetType(), "NextMonthDebug", 
                    $"console.log('Moving from month {oldMonth} to month {CurrentDate.Month} ({CurrentDate.ToString("MMMM yyyy")})');", true);
                
                // Force clear any cached data
                _events.Clear();
                    
                // Reload events for the new month
                LoadEvents();
                
                // Update the UI
                DisplayCalendar();
                
                // Force update the panel to ensure all UI elements are refreshed
                if (UpdatePanel1 != null)
                {
                    UpdatePanel1.Update();
                }
            }
            catch (Exception ex)
            {
                // Log any errors
                System.Diagnostics.Debug.WriteLine($"Error in btnNextMonth_Click: {ex.Message}");
                System.Web.UI.ScriptManager.RegisterStartupScript(this, this.GetType(), "NextMonthError", 
                    $"console.error('Error navigating to next month: {ex.Message.Replace("'", "\\'")}')", true);
            }
        }
        
        protected void DayLink_Command(object sender, CommandEventArgs e)
        {
            if (e.CommandName == "SelectDay" && e.CommandArgument != null)
            {
                if (DateTime.TryParse(e.CommandArgument.ToString(), out DateTime selectedDate))
                {
                    // Update the selected date
                    SelectedDate = selectedDate;
                    
                    // Force reload events to ensure we have the latest data
                    LoadEvents();
                    
                    // Update the calendar display
                    DisplayCalendar();
                    
                    // Add a debug message to confirm the date was selected
                    System.Diagnostics.Debug.WriteLine($"Selected date: {SelectedDate.ToString("yyyy-MM-dd")}. Events count: {_events.Count}");
                }
            }
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
                    // Add a debug message to the page
                    System.Web.UI.ScriptManager.RegisterStartupScript(this, this.GetType(), "ListNameEmpty", 
                        "alert('List name is empty. Please configure the list name in web part properties.');", true);
                    return;
                }

                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                using (SPWeb web = site.OpenWeb())
                {
                    // Check if the list exists
                    SPList list = web.Lists.TryGetList(ListName);
                    if (list == null)
                    {
                        System.Web.UI.ScriptManager.RegisterStartupScript(this, this.GetType(), "ListNotFound", 
                            $"alert('List \"{ListName}\" not found. Please check the list name.');", true);
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

                    // Try a simpler CAML query first to get all events
                    string camlQuery = $@"
                    <OrderBy>
                        <FieldRef Name='StartDate' Ascending='True' />
                    </OrderBy>";

                    SPQuery query = new SPQuery();
                    query.Query = camlQuery;
                    query.ViewFields = "<FieldRef Name='Title' /><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><FieldRef Name='Location' /><FieldRef Name='Description' /><FieldRef Name='DescriptionAr' /><FieldRef Name='Body' /><FieldRef Name='Content' /><FieldRef Name='Text' /><FieldRef Name='Comments' /><FieldRef Name='Details' /><FieldRef Name='Category' />";
                    query.ViewFieldsOnly = true;

                    query.Query = ApplyApprovedItemsFilter(list, query.Query);

                    SPListItemCollection items = list.GetItems(query);
                    
                    // Debug item count
                    if (items.Count == 0)
                    {
                        System.Web.UI.ScriptManager.RegisterStartupScript(this, this.GetType(), "NoItems", 
                            $"alert('No events found in list \"{ListName}\" for the current date range.');", true);
                    }
                    else
                    {
                        System.Web.UI.ScriptManager.RegisterStartupScript(this, this.GetType(), "ItemCount", 
                            $"console.log('Found {items.Count} events in list \"{ListName}\"');", true);
                    }

                    foreach (SPListItem item in items)
                    {
                        try
                        {
                            CalendarEvent calEvent = new CalendarEvent();
                            calEvent.Id = item.ID;
                            
                            // Debug: List all field names in this item
                            System.Text.StringBuilder fieldNames = new System.Text.StringBuilder("Available fields: ");
                            foreach (SPField field in item.Fields)
                            {
                                fieldNames.Append(field.InternalName + ", ");
                            }
                            System.Diagnostics.Debug.WriteLine(fieldNames.ToString());
                            
                            // Get Title
                            try { calEvent.Title = item["Title"] != null ? item["Title"].ToString() : string.Empty; }
                            catch { calEvent.Title = "[No Title]"; }

                            // Get StartDate
                            try
                            {
                                if (item["StartDate"] != null && DateTime.TryParse(item["StartDate"].ToString(), out DateTime startDate))
                                    calEvent.StartDate = startDate;
                                else
                                    calEvent.StartDate = DateTime.Today; // Default to today if no start date
                            }
                            catch { calEvent.StartDate = DateTime.Today; }

                            // Get EndDate
                            try
                            {
                                if (item["EndDate"] != null && DateTime.TryParse(item["EndDate"].ToString(), out DateTime endDate))
                                    calEvent.EndDate = endDate;
                                else
                                    calEvent.EndDate = calEvent.StartDate; // Default to same as start date
                            }
                            catch { calEvent.EndDate = calEvent.StartDate; }

                            // Get Location
                            try { calEvent.Location = item["Location"] != null ? item["Location"].ToString() : string.Empty; }
                            catch { calEvent.Location = string.Empty; }

                            // Get Description
                            try 
                            { 
                                if (item["Description"] != null) 
                                {
                                    calEvent.Description = item["Description"].ToString();
                                    System.Diagnostics.Debug.WriteLine($"Description for event {calEvent.Title}: {calEvent.Description}");
                                } 
                                else 
                                {
                                    // Try DescriptionAr field first since we saw it in the screenshot
                                    try {
                                        if (item["DescriptionAr"] != null) {
                                            calEvent.Description = item["DescriptionAr"].ToString();
                                            System.Diagnostics.Debug.WriteLine($"Found description in DescriptionAr field: {calEvent.Description}");
                                        }
                                    } catch {}
                                    
                                    // Try alternative field names that might contain description
                                    string[] possibleDescriptionFields = new string[] { "Body", "Content", "Text", "Comments", "Details" };
                                    foreach (string fieldName in possibleDescriptionFields)
                                    {
                                        try 
                                        {
                                            if (item.Fields.ContainsField(fieldName) && item[fieldName] != null)
                                            {
                                                calEvent.Description = item[fieldName].ToString();
                                                System.Diagnostics.Debug.WriteLine($"Found alternative description in {fieldName} field");
                                                break;
                                            }
                                        } 
                                        catch {}
                                    }
                                    
                                    if (string.IsNullOrEmpty(calEvent.Description))
                                    {
                                        // As a last resort, use the title as the description
                                        if (!string.IsNullOrEmpty(calEvent.Title))
                                        {
                                            calEvent.Description = calEvent.Title;
                                            System.Diagnostics.Debug.WriteLine($"Using title as description: {calEvent.Description}");
                                        }
                                        else
                                        {
                                            calEvent.Description = "No description available";
                                            System.Diagnostics.Debug.WriteLine($"No description found for event {calEvent.Id}");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex) 
                            { 
                                calEvent.Description = string.Empty; 
                                System.Diagnostics.Debug.WriteLine($"Error getting description: {ex.Message}");
                            }

                            // Get Category
                            try
                            {
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
                            }
                            catch { calEvent.Category = "Default"; }
                            
                            // Add the event to our collection
                            _events.Add(calEvent);
                        }
                        catch (Exception itemEx)
                        {
                            // Log individual item processing errors but continue with other items
                            System.Diagnostics.Debug.WriteLine($"Error processing item {item.ID}: {itemEx.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Display error message
                System.Web.UI.ScriptManager.RegisterStartupScript(this, this.GetType(), "ErrorMessage", 
                    $"alert('Error loading events: {ex.Message.Replace("'", "\\'")}')", true);
            }
        }

        private void DisplayCalendar()
        {
            // Force update the UI
            if (UpdatePanel1 != null)
                UpdatePanel1.Update();
            
            // Update the month name in the header
            string[] monthNames = { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
            
            // Update the month name in the navigation bar
            if (litCurrentMonth != null)
            {
                litCurrentMonth.Text = monthNames[CurrentDate.Month - 1];
            }
            
            // Update the month name and year in the top display
            if (litMonthNameDisplay != null)
            {
                litMonthNameDisplay.Text = monthNames[CurrentDate.Month - 1];
            }
            
            if (litYearDisplay != null)
            {
                litYearDisplay.Text = CurrentDate.Year.ToString();
            }
            
            // Update the month number (1-12)
            if (litMonthNumber != null)
            {
                litMonthNumber.Text = CurrentDate.Month.ToString();
            }
            
            // Data bind the controls
            Page.DataBind();
            
            // Create a list to hold the calendar days
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

            // Check if there are any events for the selected date
            if (selectedDateEvents.Count > 0)
            {
                rptEvents.DataSource = selectedDateEvents;
                rptEvents.DataBind();
                pnlNoEvents.Visible = false;
            }
            else
            {
                rptEvents.DataSource = null;
                rptEvents.DataBind();
                pnlNoEvents.Visible = true;
            }
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
                int maxEventsToShow = 2; // Reduced to 2 to prevent overflow
                
                // Get events for this day
                List<CalendarEvent> eventsToShow = day.Events.Take(maxEventsToShow).ToList();
                
                // If there are more events than we can show, add a count indicator
                if (day.Events.Count > maxEventsToShow)
                {
                    // We'll handle this in the UI if needed
                    System.Diagnostics.Debug.WriteLine($"Day {day.Date.ToString("yyyy-MM-dd")} has {day.Events.Count} events, showing {maxEventsToShow}");
                }

                rptDayEvents.DataSource = eventsToShow;
                rptDayEvents.DataBind();
            }
        }

        protected void rptEvents_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {
            // Additional event binding logic can be added here if needed
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
            try
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
                    { "Personal", "#FB8C00" },  // Orange
                    { "Default", "#3174ad" }    // Medium Blue
                };

                if (!string.IsNullOrEmpty(calEvent.Category) && categoryColors.ContainsKey(calEvent.Category))
                    return categoryColors[calEvent.Category];

                // Assign colors based on the first letter of the event title if no category
                if (!string.IsNullOrEmpty(calEvent.Title))
                {
                    char firstChar = char.ToUpper(calEvent.Title[0]);
                    string[] colorArray = new string[] { 
                        "#4285F4", "#EA4335", "#FBBC05", "#34A853", "#8E24AA", 
                        "#FB8C00", "#0097A7", "#D81B60", "#5E35B1", "#1E88E5" 
                    };
                    int index = (int)firstChar % colorArray.Length;
                    return colorArray[index];
                }

                // Default color if all else fails
                return "#3174ad";
            }
            catch
            {
                // Fallback color if any error occurs
                return "#3174ad";
            }
        }

        
        protected string TruncateDescription(string description)
        {
            if (string.IsNullOrEmpty(description))
                return string.Empty;
                
            // Limit description to 150 characters
            int maxLength = 150;
            if (description.Length <= maxLength)
                return description;
                
            // Truncate at the last space before maxLength to avoid cutting words
            int lastSpace = description.Substring(0, maxLength).LastIndexOf(' ');
            if (lastSpace <= 0) // No space found, just cut at maxLength
                lastSpace = maxLength;
                
            return description.Substring(0, lastSpace) + "...";
        }
        
        protected string GetEventDetailsUrl(object dataItem)
        {
            try
            {
                CalendarEvent calEvent = (CalendarEvent)dataItem;
                
                string currentWebUrl = SPContext.Current.Web.Url;
                string listWebUrl = HttpContext.Current?.Request?.Url?.GetLeftPart(UriPartial.Authority);
                
                // Construct the URL to the event details page with the event ID as a query string parameter
                // This assumes there's a page called EventDetails.aspx in the same site that hosts the EventDetailsWP web part
                string detailsUrl = $"{currentWebUrl}/EventDetails.aspx?EventID={calEvent.Id}";
                
                // Alternatively, if you want to navigate to a specific page in the Pages library:
                // string detailsUrl = $"{currentUrl}/Pages/EventDetails.aspx?EventID={calEvent.Id}";
                
                // You can also add additional parameters if needed
                // For example, to pass the list name:
                if (!string.IsNullOrEmpty(ListName))
                {
                    detailsUrl += $"&ListName={SPEncode.UrlEncode(ListName)}";
                }

                if (!string.IsNullOrEmpty(listWebUrl))
                {
                    detailsUrl += $"&ListWebUrl={SPEncode.UrlEncode(listWebUrl)}";
                }
                
                return detailsUrl;
            }
            catch
            {
                // Return a safe fallback URL if there's an error
                return "#";
            }
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
