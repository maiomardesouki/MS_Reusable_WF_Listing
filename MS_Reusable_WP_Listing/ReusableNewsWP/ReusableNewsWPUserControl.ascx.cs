using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;

namespace MS_Reusable_WP_Listing.ReusableNewsWP
{
    public partial class ReusableNewsWPUserControl : UserControl
    {
        // Properties passed from the web part
        public string ListName { get; set; }
        public string WebUrl { get; set; }
        public string TitleField { get; set; }
        public string DescriptionField { get; set; }
        public string DateField { get; set; }
        public string LocationField { get; set; }
        public string ImageField { get; set; }
        public string DetailsPageUrl { get; set; }
        public int PageSize { get; set; }

        // For timeout detection
        private Stopwatch stopwatch;
        private const int MaxExecutionTimeMs = 10000; // 10 seconds

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                // Set the page size for the data pager
                DataPager pager = lsvNews.FindControl("_dpPager") as DataPager;
                if (pager != null)
                {
                    pager.PageSize = PageSize;
                }
                
                LoadItems();
            }
        }

        private void LoadItems()
        {
            stopwatch = Stopwatch.StartNew();
            pnlContent.Visible = true;
            pnlError.Visible = false;

            try
            {
                // Check if list name is provided
                if (string.IsNullOrEmpty(ListName))
                {
                    ShowError("Please configure the list name in the web part properties.");
                    return;
                }

                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                using (SPWeb web = string.IsNullOrEmpty(WebUrl) ? site.OpenWeb() : site.OpenWeb(WebUrl))
                {
                    // Check if the list exists
                    Debug.WriteLine($"Looking for list: '{ListName}' in web: '{web.Url}'");
                    SPList list = web.Lists.TryGetList(ListName);
                    if (list == null)
                    {
                        ShowError($"The list '{ListName}' does not exist.");
                        // Log available lists for debugging
                        StringBuilder listsLog = new StringBuilder();
                        listsLog.AppendLine("Available lists:");
                        foreach (SPList availableList in web.Lists)
                        {
                            listsLog.AppendLine($" - {availableList.Title}");
                        }
                        Debug.WriteLine(listsLog.ToString());
                        return;
                    }
                    
                    Debug.WriteLine($"Found list: '{list.Title}' with {list.ItemCount} items");

                    // Build the ViewFields XML
                    StringBuilder viewFieldsBuilder = new StringBuilder();
                    viewFieldsBuilder.Append("<FieldRef Name='ID' />");
                    
                    // Add title field
                    if (!string.IsNullOrEmpty(TitleField))
                        viewFieldsBuilder.Append($"<FieldRef Name='{TitleField}' />");
                    
                    // Add description field
                    if (!string.IsNullOrEmpty(DescriptionField))
                        viewFieldsBuilder.Append($"<FieldRef Name='{DescriptionField}' />");
                    
                    // Add date field
                    if (!string.IsNullOrEmpty(DateField))
                        viewFieldsBuilder.Append($"<FieldRef Name='{DateField}' />");
                    
                    // Add location field
                    if (!string.IsNullOrEmpty(LocationField))
                        viewFieldsBuilder.Append($"<FieldRef Name='{LocationField}' />");
                    
                    // Add image field
                    if (!string.IsNullOrEmpty(ImageField))
                        viewFieldsBuilder.Append($"<FieldRef Name='{ImageField}' />");

                    SPQuery query = new SPQuery();
                    query.ViewFields = viewFieldsBuilder.ToString();
                    query.ViewFieldsOnly = true;
                    query.RowLimit = (uint)PageSize;
                    
                    // Only add OrderBy if DateField is specified
                    if (!string.IsNullOrEmpty(DateField))
                    {
                        query.Query = "<OrderBy><FieldRef Name='" + DateField + "' Ascending='False' /></OrderBy>";
                        Debug.WriteLine($"Ordering by field: '{DateField}'");
                    }
                    else
                    {
                        Debug.WriteLine("No date field specified for ordering");
                    }

                    query.Query = ApplyApprovedItemsFilter(list, query.Query);
                    
                    Debug.WriteLine($"Query ViewFields: {query.ViewFields}");

                    
                    SPListItemCollection items = list.GetItems(query);
                    Debug.WriteLine($"Query returned {items.Count} items");

                    if (items.Count == 0)
                    {
                        ShowError("No items found in the list.", false);
                        return;
                    }

                    // Create a DataTable to hold the items with additional properties
                    DataTable dt = new DataTable();
                    dt.Columns.Add("ID", typeof(int));
                    dt.Columns.Add("Title", typeof(string));
                    dt.Columns.Add("Description", typeof(string));
                    dt.Columns.Add("DateDay", typeof(string));
                    dt.Columns.Add("DateMonth", typeof(string));
                    dt.Columns.Add("Location", typeof(string));
                    dt.Columns.Add("ImageUrl", typeof(string));
                    dt.Columns.Add("DetailsUrl", typeof(string));

                    // Log configured field names
                    Debug.WriteLine($"Configured fields - Title: '{TitleField}', Description: '{DescriptionField}', Date: '{DateField}', Location: '{LocationField}', Image: '{ImageField}'");
                    
                    // Check if the configured fields exist in the list
                    Debug.WriteLine("Checking if configured fields exist in the list:");
                    if (!string.IsNullOrEmpty(TitleField))
                        Debug.WriteLine($"Title field '{TitleField}' exists: {list.Fields.ContainsField(TitleField)}");
                    if (!string.IsNullOrEmpty(DescriptionField))
                        Debug.WriteLine($"Description field '{DescriptionField}' exists: {list.Fields.ContainsField(DescriptionField)}");
                    if (!string.IsNullOrEmpty(DateField))
                        Debug.WriteLine($"Date field '{DateField}' exists: {list.Fields.ContainsField(DateField)}");
                    if (!string.IsNullOrEmpty(LocationField))
                        Debug.WriteLine($"Location field '{LocationField}' exists: {list.Fields.ContainsField(LocationField)}");
                    if (!string.IsNullOrEmpty(ImageField))
                        Debug.WriteLine($"Image field '{ImageField}' exists: {list.Fields.ContainsField(ImageField)}");
                    
                    // Fill the DataTable with items and computed properties
                    foreach (SPListItem item in items)
                    {
                        try
                        {
                            Debug.WriteLine($"Processing item ID: {item.ID}");
                            DataRow row = dt.NewRow();
                            row["ID"] = item.ID;
                            
                            string titleValue = GetTitle(item);
                            Debug.WriteLine($"Title: {titleValue}");
                            row["Title"] = titleValue;
                            
                            string descValue = GetDescription(item);
                            Debug.WriteLine($"Description: {(descValue.Length > 50 ? descValue.Substring(0, 50) + "..." : descValue)}");
                            row["Description"] = descValue;
                            
                            string dayValue = FormatDateDay(item);
                            Debug.WriteLine($"DateDay: {dayValue}");
                            row["DateDay"] = dayValue;
                            
                            string monthValue = FormatDateMonth(item);
                            Debug.WriteLine($"DateMonth: {monthValue}");
                            row["DateMonth"] = monthValue;
                            
                            string locationValue = GetLocation(item);
                            Debug.WriteLine($"Location: {locationValue}");
                            row["Location"] = locationValue;

                            // Get image URL with special handling for publishing images
                            string imageUrl = GetImageUrl(item);
                            Debug.WriteLine($"Image URL for item {item.ID}: {imageUrl}");
                            
                            // Additional debug info for image fields
                            if (item.Fields.ContainsField(ImageField) && item[ImageField] != null)
                            {
                                SPField fieldObj = item.Fields.GetFieldByInternalName(ImageField);
                                Debug.WriteLine($"Field type: {fieldObj.GetType().Name}, Field value: {item[ImageField]}");
                            }
                            
                            row["ImageUrl"] = imageUrl;

                            string detailsUrl = GetDetailsUrl(item);
                            Debug.WriteLine($"DetailsUrl: {detailsUrl}");
                            row["DetailsUrl"] = detailsUrl;
                            
                            dt.Rows.Add(row);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error processing item {item.ID}: {ex.Message}");
                            Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                        }
                    }

                    // Bind the DataTable to the ListView instead of the SPListItemCollection
                    Debug.WriteLine($"Binding DataTable with {dt.Rows.Count} rows to ListView");
                    
                    if (dt.Rows.Count > 0)
                    {
                        lsvNews.DataSource = dt;
                        lsvNews.DataBind();
                        
                        // Ensure content panel is visible and error panel is hidden
                        pnlContent.Visible = true;
                        pnlError.Visible = false;
                        
                        Debug.WriteLine("DataTable bound to ListView successfully");
                    }
                    else
                    {
                        // No rows in DataTable, show empty message
                        Debug.WriteLine("No rows in DataTable, showing empty message");
                        ShowError("No news items found with the configured fields.", false);
                    }
                    
                    // Check for timeout
                    if (stopwatch.ElapsedMilliseconds > MaxExecutionTimeMs)
                    {
                        ShowError("The operation timed out. Please try again or contact your administrator.");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error loading news items: {ex.Message}");
                Debug.WriteLine($"Error in ReusableNewsWP: {ex.ToString()}");
            }
            finally
            {
                stopwatch.Stop();
            }
        }

        protected void lsvNews_PagePropertiesChanging(object sender, PagePropertiesChangingEventArgs e)
        {
            // Get the data pager
            DataPager pager = lsvNews.FindControl("_dpPager") as DataPager;
            if (pager != null)
            {
                pager.SetPageProperties(e.StartRowIndex, e.MaximumRows, false);
                LoadItems();
            }
        }

        private void ShowError(string message, bool isError = true)
        {
            Debug.WriteLine($"ShowError called with message: {message}, isError: {isError}");
            litError.Text = message;
            pnlError.Visible = true;
            pnlContent.Visible = !isError;
            pnlError.CssClass = isError ? "error-message" : "info-message";
            
            // Don't try to update the panel during initial page load
            // This avoids the UpdatePanel.Update() error
        }

        private static string ApplyApprovedItemsFilter(SPList list, string existingCaml)
        {
            if (list == null || !list.EnableModeration)
                return existingCaml;

            string moderationWhere = "<Eq><FieldRef Name='_ModerationStatus' /><Value Type='ModStat'>0</Value></Eq>";

            if (string.IsNullOrEmpty(existingCaml))
                return "<Where>" + moderationWhere + "</Where>";

            Regex whereRegex = new Regex("<Where>(.*?)</Where>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match whereMatch = whereRegex.Match(existingCaml);
            if (whereMatch.Success)
            {
                string innerWhere = whereMatch.Groups[1].Value;
                string newWhere = "<Where><And>" + innerWhere + moderationWhere + "</And></Where>";
                return whereRegex.Replace(existingCaml, newWhere, 1);
            }

            return "<Where>" + moderationWhere + "</Where>" + existingCaml;
        }

        // Helper methods for formatting data
        public string GetTitle(object dataItem)
        {
            SPListItem item = dataItem as SPListItem;
            if (item == null)
                return "No Item";
                
            if (string.IsNullOrEmpty(TitleField))
                return "Title Field Not Configured";

            try
            {
                if (item.Fields.ContainsField(TitleField))
                {
                    if (item[TitleField] != null)
                    {
                        return item[TitleField].ToString();
                    }
                    else
                    {
                        Debug.WriteLine($"Title field '{TitleField}' is null for item {item.ID}");
                        return "[No Title]";
                    }
                }
                else
                {
                    Debug.WriteLine($"Title field '{TitleField}' does not exist in the list");
                    return "[Field Not Found]";
                }
            }
            catch (Exception ex) 
            { 
                Debug.WriteLine($"Error getting title: {ex.Message}");
                return "[Error]";
            }
        }

        public string GetDescription(object dataItem)
        {
            SPListItem item = dataItem as SPListItem;
            if (item == null)
                return string.Empty;
                
            if (string.IsNullOrEmpty(DescriptionField))
                return string.Empty;

            try
            {
                if (item.Fields.ContainsField(DescriptionField))
                {
                    if (item[DescriptionField] != null)
                    {
                        string desc = item[DescriptionField].ToString();
                        // Strip HTML if present
                        if (desc.Contains("<") && desc.Contains(">"))
                        {
                            desc = Regex.Replace(desc, @"<[^>]+>", string.Empty);
                        }
                        return desc;
                    }
                    else
                    {
                        Debug.WriteLine($"Description field '{DescriptionField}' is null for item {item.ID}");
                        return string.Empty;
                    }
                }
                else
                {
                    Debug.WriteLine($"Description field '{DescriptionField}' does not exist in the list");
                    return string.Empty;
                }
            }
            catch (Exception ex) 
            { 
                Debug.WriteLine($"Error getting description: {ex.Message}");
                return string.Empty;
            }
        }

        public string FormatDateDay(object dataItem)
        {
            SPListItem item = dataItem as SPListItem;
            if (item == null)
                return string.Empty;
                
            if (string.IsNullOrEmpty(DateField))
                return string.Empty;

            try
            {
                if (item.Fields.ContainsField(DateField))
                {
                    object value = item[DateField];
                    if (value != null)
                    {
                        // Try to parse as DateTime
                        if (value is DateTime)
                        {
                            DateTime date = (DateTime)value;
                            return date.Day.ToString();
                        }
                        else if (DateTime.TryParse(value.ToString(), out DateTime parsedDate))
                        {
                            return parsedDate.Day.ToString();
                        }
                        else
                        {
                            Debug.WriteLine($"Date field '{DateField}' value '{value}' could not be parsed as date for item {item.ID}");
                            return string.Empty;
                        }
                    }
                    else
                    {
                        Debug.WriteLine($"Date field '{DateField}' is null for item {item.ID}");
                        return string.Empty;
                    }
                }
                else
                {
                    Debug.WriteLine($"Date field '{DateField}' does not exist in the list");
                    return string.Empty;
                }
            }
            catch (Exception ex) 
            { 
                Debug.WriteLine($"Error getting date day: {ex.Message}");
                return string.Empty;
            }
        }

        public string FormatDateMonth(object dataItem)
        {
            SPListItem item = dataItem as SPListItem;
            if (item == null)
                return string.Empty;
                
            if (string.IsNullOrEmpty(DateField))
                return string.Empty;

            try
            {
                if (item.Fields.ContainsField(DateField))
                {
                    object value = item[DateField];
                    if (value != null)
                    {
                        // Try to parse as DateTime
                        if (value is DateTime)
                        {
                            DateTime date = (DateTime)value;
                            return date.ToString("MMMM", CultureInfo.InvariantCulture).ToUpper();
                        }
                        else if (DateTime.TryParse(value.ToString(), out DateTime parsedDate))
                        {
                            return parsedDate.ToString("MMMM", CultureInfo.InvariantCulture).ToUpper();
                        }
                        else
                        {
                            Debug.WriteLine($"Date field '{DateField}' value '{value}' could not be parsed as date for item {item.ID}");
                            return string.Empty;
                        }
                    }
                    else
                    {
                        Debug.WriteLine($"Date field '{DateField}' is null for item {item.ID}");
                        return string.Empty;
                    }
                }
                else
                {
                    Debug.WriteLine($"Date field '{DateField}' does not exist in the list");
                    return string.Empty;
                }
            }
            catch (Exception ex) 
            { 
                Debug.WriteLine($"Error getting date month: {ex.Message}");
                return string.Empty;
            }
        }

        public string GetLocation(object dataItem)
        {
            SPListItem item = dataItem as SPListItem;
            if (item == null)
                return string.Empty;
                
            if (string.IsNullOrEmpty(LocationField))
                return string.Empty;

            try
            {
                if (item.Fields.ContainsField(LocationField))
                {
                    if (item[LocationField] != null)
                    {
                        return item[LocationField].ToString();
                    }
                    else
                    {
                        Debug.WriteLine($"Location field '{LocationField}' is null for item {item.ID}");
                        return string.Empty;
                    }
                }
                else
                {
                    Debug.WriteLine($"Location field '{LocationField}' does not exist in the list");
                    return string.Empty;
                }
            }
            catch (Exception ex) 
            { 
                Debug.WriteLine($"Error getting location: {ex.Message}");
                return string.Empty;
            }
        }

        // Method to get the image URL from the specified field
        public string GetImageUrl(object dataItem)
        {
            SPListItem item = dataItem as SPListItem;
            return ImageHelper.GetImageUrl(item, ImageField);
        }

        public string GetDetailsUrl(object dataItem)
        {
            SPListItem item = dataItem as SPListItem;
            if (item == null)
                return "#";

            return $"{DetailsPageUrl}?aID={item.ID}";
        }
    }
}
