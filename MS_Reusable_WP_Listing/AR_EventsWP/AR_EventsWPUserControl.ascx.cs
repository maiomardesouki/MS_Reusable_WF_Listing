using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;

namespace MS_Reusable_WP_Listing.AR_EventsWP
{
    public partial class AR_EventsWPUserControl : UserControl
    {
        private string _listName = string.Empty;
        public string ListName
        {
            get { return _listName; }
            set { _listName = value; }
        }

        private string _calendarTitle = "التقويم";
        public string CalendarTitle
        {
            get { return _calendarTitle; }
            set { _calendarTitle = value; }
        }

        public DateTime CurrentDate
        {
            get { return ViewState["CurrentDate"] as DateTime? ?? DateTime.Today; }
            set { ViewState["CurrentDate"] = value; }
        }

        public DateTime SelectedDate
        {
            get { return ViewState["SelectedDate"] as DateTime? ?? DateTime.Today; }
            set { ViewState["SelectedDate"] = value; }
        }

        private readonly List<CalendarEvent> _events = new List<CalendarEvent>();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                CurrentDate = DateTime.Today;
                SelectedDate = DateTime.Today;
                LoadEvents();
                DisplayCalendar();
            }
        }

        protected void btnPrevMonth_Click(object sender, EventArgs e)
        {
            try
            {
                CurrentDate = CurrentDate.AddMonths(-1);
                SelectedDate = new DateTime(CurrentDate.Year, CurrentDate.Month, 1);
                _events.Clear();
                LoadEvents();
                DisplayCalendar();
                if (UpdatePanel1 != null)
                    UpdatePanel1.Update();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in btnPrevMonth_Click: {ex.Message}");
            }
        }

        protected void btnNextMonth_Click(object sender, EventArgs e)
        {
            try
            {
                CurrentDate = CurrentDate.AddMonths(1);
                SelectedDate = new DateTime(CurrentDate.Year, CurrentDate.Month, 1);
                _events.Clear();
                LoadEvents();
                DisplayCalendar();
                if (UpdatePanel1 != null)
                    UpdatePanel1.Update();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in btnNextMonth_Click: {ex.Message}");
            }
        }

        protected void DayLink_Command(object sender, CommandEventArgs e)
        {
            if (e.CommandName == "SelectDay" && e.CommandArgument != null)
            {
                if (DateTime.TryParse(e.CommandArgument.ToString(), out DateTime selectedDate))
                {
                    SelectedDate = selectedDate;
                    LoadEvents();
                    DisplayCalendar();
                }
            }
        }

        private void LoadEvents()
        {
            _events.Clear();

            try
            {
                if (string.IsNullOrEmpty(ListName))
                {
                    System.Web.UI.ScriptManager.RegisterStartupScript(this, this.GetType(), "ListNameEmpty",
                        "alert('اسم القائمة فارغ. يرجى تحديد اسم القائمة من خصائص الويب بارت.');", true);
                    return;
                }

                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                using (SPWeb web = site.OpenWeb())
                {
                    SPList list = web.Lists.TryGetList(ListName);
                    if (list == null)
                    {
                        System.Web.UI.ScriptManager.RegisterStartupScript(this, this.GetType(), "ListNotFound",
                            $"alert('القائمة \"{ListName}\" غير موجودة.');", true);
                        return;
                    }

                    string titleArInternalName = ResolveFieldInternalName(list, "TitleAR");
                    bool hasTitleArField = !string.IsNullOrEmpty(titleArInternalName);

                    string camlQuery = @"
                    <OrderBy>
                        <FieldRef Name='StartDate' Ascending='True' />
                    </OrderBy>";

                    SPQuery query = new SPQuery();
                    query.Query = camlQuery;
                    query.ViewFields = (hasTitleArField ? $"<FieldRef Name='{titleArInternalName}' />" : "") + "<FieldRef Name='Title' /><FieldRef Name='StartDate' /><FieldRef Name='EndDate' /><FieldRef Name='Location' /><FieldRef Name='Description' /><FieldRef Name='DescriptionAr' /><FieldRef Name='Body' /><FieldRef Name='Content' /><FieldRef Name='Text' /><FieldRef Name='Comments' /><FieldRef Name='Details' /><FieldRef Name='Category' />";
                    query.ViewFieldsOnly = true;

                    query.Query = ApplyApprovedItemsFilter(list, query.Query);

                    SPListItemCollection items = list.GetItems(query);

                    foreach (SPListItem item in items)
                    {
                        try
                        {
                            CalendarEvent calEvent = new CalendarEvent();
                            calEvent.Id = item.ID;

                            try
                            {
                                if (hasTitleArField)
                                {
                                    calEvent.Title = item[titleArInternalName] != null ? item[titleArInternalName].ToString() : string.Empty;
                                }
                                else
                                {
                                    calEvent.Title = item["Title"] != null ? item["Title"].ToString() : string.Empty;
                                }
                            }
                            catch
                            {
                                calEvent.Title = string.Empty;
                            }

                            try
                            {
                                if (item["StartDate"] != null && DateTime.TryParse(item["StartDate"].ToString(), out DateTime startDate))
                                    calEvent.StartDate = startDate;
                                else
                                    calEvent.StartDate = DateTime.Today;
                            }
                            catch { calEvent.StartDate = DateTime.Today; }

                            try
                            {
                                if (item["EndDate"] != null && DateTime.TryParse(item["EndDate"].ToString(), out DateTime endDate))
                                    calEvent.EndDate = endDate;
                                else
                                    calEvent.EndDate = calEvent.StartDate;
                            }
                            catch { calEvent.EndDate = calEvent.StartDate; }

                            try { calEvent.Location = item["Location"] != null ? item["Location"].ToString() : string.Empty; }
                            catch { calEvent.Location = string.Empty; }

                            try
                            {
                                if (item["DescriptionAr"] != null)
                                    calEvent.Description = item["DescriptionAr"].ToString();
                                else if (item["Description"] != null)
                                    calEvent.Description = item["Description"].ToString();
                                else
                                    calEvent.Description = string.Empty;

                                if (string.IsNullOrEmpty(calEvent.Description) && !string.IsNullOrEmpty(calEvent.Title))
                                    calEvent.Description = calEvent.Title;
                            }
                            catch { calEvent.Description = string.Empty; }

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
                                else
                                {
                                    calEvent.Category = "Default";
                                }
                            }
                            catch { calEvent.Category = "Default"; }

                            _events.Add(calEvent);
                        }
                        catch (Exception itemEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error processing item {item.ID}: {itemEx.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Web.UI.ScriptManager.RegisterStartupScript(this, this.GetType(), "ErrorMessage",
                    $"alert('حدث خطأ أثناء تحميل الفعاليات: {ex.Message.Replace("'", "\\'")}')", true);
            }
        }

        private void DisplayCalendar()
        {
            if (UpdatePanel1 != null)
                UpdatePanel1.Update();

            string monthName = GetArabicMonthName(CurrentDate.Month);

            if (litCurrentMonth != null)
                litCurrentMonth.Text = monthName;

            if (litMonthNameDisplay != null)
                litMonthNameDisplay.Text = monthName;

            if (litYearDisplay != null)
                litYearDisplay.Text = ToArabicIndicDigits(CurrentDate.Year.ToString());

            if (litMonthNumber != null)
                litMonthNumber.Text = ToArabicIndicDigits(CurrentDate.Month.ToString());

            Page.DataBind();

            List<CalendarDay> calendarDays = new List<CalendarDay>();
            DateTime firstDayOfMonth = new DateTime(CurrentDate.Year, CurrentDate.Month, 1);
            int daysToSubtract = (int)firstDayOfMonth.DayOfWeek;
            DateTime firstDayOfCalendarGrid = firstDayOfMonth.AddDays(-daysToSubtract);

            for (int i = 0; i < 42; i++)
            {
                DateTime currentDay = firstDayOfCalendarGrid.AddDays(i);
                CalendarDay calendarDay = new CalendarDay();
                calendarDay.Date = currentDay;
                calendarDay.IsCurrentMonth = currentDay.Month == CurrentDate.Month;
                calendarDay.IsToday = currentDay.Date == DateTime.Today.Date;
                calendarDay.IsSelected = currentDay.Date == SelectedDate.Date;
                calendarDay.Events = _events.Where(ev => ev.StartDate.Date <= currentDay.Date && ev.EndDate.Date >= currentDay.Date).ToList();
                calendarDays.Add(calendarDay);
            }

            rptCalendarDays.DataSource = calendarDays;
            rptCalendarDays.DataBind();

            List<CalendarEvent> selectedDateEvents = _events.Where(ev => ev.StartDate.Date <= SelectedDate.Date && ev.EndDate.Date >= SelectedDate.Date).ToList();

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

        protected void rptCalendarDays_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {
            if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
            {
                CalendarDay day = (CalendarDay)e.Item.DataItem;
                Repeater rptDayEvents = (Repeater)e.Item.FindControl("rptDayEvents");

                int maxEventsToShow = 2;
                List<CalendarEvent> eventsToShow = day.Events.Take(maxEventsToShow).ToList();
                rptDayEvents.DataSource = eventsToShow;
                rptDayEvents.DataBind();
            }
        }

        protected void rptEvents_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {
        }

        public string GetCalendarTitle()
        {
            return string.IsNullOrEmpty(CalendarTitle) ? "التقويم" : CalendarTitle;
        }

        public string FormatArabicDay(object rawDate)
        {
            DateTime? dt = CoerceDate(rawDate);
            if (!dt.HasValue)
                return string.Empty;
            return ToArabicIndicDigits(dt.Value.Day.ToString("00"));
        }

        public string FormatArabicMonthAbbrev(object rawDate)
        {
            DateTime? dt = CoerceDate(rawDate);
            if (!dt.HasValue)
                return string.Empty;
            return GetArabicMonthName(dt.Value.Month);
        }

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
            return ToArabicIndicDigits(day.Date.Day.ToString());
        }

        protected string GetEventColor(object dataItem)
        {
            try
            {
                CalendarEvent calEvent = (CalendarEvent)dataItem;

                Dictionary<string, string> categoryColors = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "Meeting", "#4285F4" },
                    { "Holiday", "#EA4335" },
                    { "Training", "#FBBC05" },
                    { "Travel", "#34A853" },
                    { "Conference", "#8E24AA" },
                    { "Personal", "#FB8C00" },
                    { "Default", "#3174ad" }
                };

                if (!string.IsNullOrEmpty(calEvent.Category) && categoryColors.ContainsKey(calEvent.Category))
                    return categoryColors[calEvent.Category];

                if (!string.IsNullOrEmpty(calEvent.Title))
                {
                    char firstChar = char.ToUpper(calEvent.Title[0]);
                    string[] colorArray = new string[]
                    {
                        "#4285F4", "#EA4335", "#FBBC05", "#34A853", "#8E24AA",
                        "#FB8C00", "#0097A7", "#D81B60", "#5E35B1", "#1E88E5"
                    };
                    int index = (int)firstChar % colorArray.Length;
                    return colorArray[index];
                }

                return "#3174ad";
            }
            catch
            {
                return "#3174ad";
            }
        }

        protected string TruncateDescription(string description)
        {
            if (string.IsNullOrEmpty(description))
                return string.Empty;

            int maxLength = 150;
            if (description.Length <= maxLength)
                return description;

            int lastSpace = description.Substring(0, maxLength).LastIndexOf(' ');
            if (lastSpace <= 0)
                lastSpace = maxLength;

            return description.Substring(0, lastSpace) + "...";
        }

        protected string GetEventDetailsUrl(object dataItem)
        {
            try
            {
                CalendarEvent calEvent = (CalendarEvent)dataItem;
                string currentUrl = SPContext.Current.Web.Url;
                string detailsUrl = $"{currentUrl}/EventDetails.aspx?EventID={calEvent.Id}";
                if (!string.IsNullOrEmpty(ListName))
                    detailsUrl += $"&ListName={SPEncode.UrlEncode(ListName)}";
                return detailsUrl;
            }
            catch
            {
                return "#";
            }
        }

        private static DateTime? CoerceDate(object raw)
        {
            if (raw == null)
                return null;
            if (raw is DateTime)
                return (DateTime)raw;

            DateTime parsed;
            if (DateTime.TryParse(raw.ToString(), out parsed))
                return parsed;

            return null;
        }

        private static string GetArabicMonthName(int month)
        {
            string[] months = new string[]
            {
                "يناير",
                "فبراير",
                "مارس",
                "أبريل",
                "مايو",
                "يونيو",
                "يوليو",
                "أغسطس",
                "سبتمبر",
                "أكتوبر",
                "نوفمبر",
                "ديسمبر"
            };

            if (month < 1 || month > 12)
                return string.Empty;
            return months[month - 1];
        }

        private static string ToArabicIndicDigits(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            char[] arabicDigits = new char[] { '٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩' };
            char[] chars = input.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] >= '0' && chars[i] <= '9')
                    chars[i] = arabicDigits[chars[i] - '0'];
            }
            return new string(chars);
        }

        private static string ResolveFieldInternalName(SPList list, string requestedName)
        {
            if (list == null || string.IsNullOrEmpty(requestedName))
                return null;

            foreach (SPField field in list.Fields)
            {
                if (field == null)
                    continue;

                if (string.Equals(field.InternalName, requestedName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(field.StaticName, requestedName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(field.Title, requestedName, StringComparison.OrdinalIgnoreCase))
                {
                    return field.InternalName;
                }
            }

            return null;
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
    }

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
}
