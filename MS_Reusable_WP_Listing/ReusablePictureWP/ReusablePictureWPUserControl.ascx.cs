using System;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Data;
using Microsoft.SharePoint;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Utilities;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace MS_Reusable_WP_Listing.ReusablePictureWP
{
    public partial class ReusablePictureWPUserControl : UserControl
    {
        #region Properties

        // Property to store the list name from the web part properties
        private string _listName = string.Empty;
        public string ListName
        {
            get { return _listName; }
            set { _listName = value; }
        }
        
        // Property for the web URL
        private string _webUrl = string.Empty;
        public string WebUrl
        {
            get { return _webUrl; }
            set { _webUrl = value; }
        }

        // Property for the title field name
        private string _titleField = "Title";
        public string TitleField
        {
            get { return _titleField; }
            set { _titleField = value; }
        }

        // Property for the description field name
        private string _descriptionField = "Description";
        public string DescriptionField
        {
            get { return _descriptionField; }
            set { _descriptionField = value; }
        }

        // Property for the image field name
        private string _imageField = "Picture";
        public string ImageField
        {
            get { return _imageField; }
            set { _imageField = value; }
        }

        // Property for the date field name
        private string _dateField = "Date";
        public string DateField
        {
            get { return _dateField; }
            set { _dateField = value; }
        }

        // Property for the hyperlink field name
        private string _hyperlinkField = "Link";
        public string HyperlinkField
        {
            get { return _hyperlinkField; }
            set { _hyperlinkField = value; }
        }

        // Property for the page size
        private int _pageSize = 6;
        public int PageSize
        {
            get { return _pageSize; }
            set { _pageSize = value; }
        }

        #endregion

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                // Set the page size for the data pager
                DataPager pager = lsvPictures.FindControl("_dpPager") as DataPager;
                if (pager != null)
                {
                    pager.PageSize = PageSize;
                }

                // Load items from the SharePoint list
                LoadItems();
            }
        }

        protected void lsvPictures_PagePropertiesChanging(object sender, PagePropertiesChangingEventArgs e)
        {
            // Get the data pager
            DataPager pager = lsvPictures.FindControl("_dpPager") as DataPager;
            if (pager != null)
            {
                pager.SetPageProperties(e.StartRowIndex, e.MaximumRows, false);
                LoadItems();
            }
        }
        
        // Cache for storing items to avoid repeated SharePoint queries
        private static SPListItemCollection _cachedItems = null;
        private static DateTime _cacheExpiry = DateTime.MinValue;
        private static string _cacheKey = string.Empty;
        
        private void LoadItems()
        {
            try
            {
                _cachedItems = null;
                _cacheExpiry = DateTime.MinValue;
                _cacheKey = string.Empty;
                _descriptionCache.Clear();
                _hyperlinkCache.Clear();

                // Check if list name is provided
                if (string.IsNullOrEmpty(ListName))
                {
                    ShowError("Please configure the list name in the web part properties.");
                    return;
                }
                
                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                {
                    // Get the web based on the WebUrl property
                    SPWeb web = string.IsNullOrEmpty(WebUrl) ? site.OpenWeb() : site.OpenWeb(WebUrl);

                    using (web)
                    {
                        // Get the list
                        SPList list = web.Lists.TryGetList(ListName);
                        if (list == null)
                        {
                            ShowError($"List '{ListName}' not found in web '{web.Url}'.");
                            return;
                        }

                        // Create an optimized query with only the necessary fields
                        SPQuery query = new SPQuery();
                        query.ViewFields = "<FieldRef Name='ID' />";
                        
                        // Only include fields we actually need
                        if (!string.IsNullOrEmpty(TitleField)) query.ViewFields += $"<FieldRef Name='{TitleField}' />";
                        if (!string.IsNullOrEmpty(DescriptionField)) query.ViewFields += $"<FieldRef Name='{DescriptionField}' />";
                        if (!string.IsNullOrEmpty(DateField)) 
                        {
                            query.ViewFields += $"<FieldRef Name='{DateField}' />";
                            query.Query = $"<OrderBy><FieldRef Name='{DateField}' Ascending='False' /></OrderBy>";
                        }
                        if (!string.IsNullOrEmpty(HyperlinkField)) query.ViewFields += $"<FieldRef Name='{HyperlinkField}' />";

                        query.Query = ApplyApprovedItemsFilter(list, query.Query);

                        // Get items
                        SPListItemCollection items = list.GetItems(query);

                        if (items.Count == 0)
                        {
                            ShowError("No items found in the list.", false);
                            return;
                        }
                        
                        // Bind data to the ListView
                        lsvPictures.DataSource = items;
                        lsvPictures.DataBind();
                        
                        // Show the content panel
                        pnlContent.Visible = true;
                        pnlError.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error loading items: {ex.Message}");
            }
        }
        
        private void ShowError(string message, bool isError = true)
        {
            // Log the error for debugging
            Debug.WriteLine($"{(isError ? "ERROR" : "INFO")}: {message}");
            
            // Update the UI
            litError.Text = message;
            pnlError.Visible = true;
            pnlContent.Visible = false;
            
            // Make sure the error is displayed immediately
            UpdatePanel1?.Update();
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
        
        public string GetImageUrl(object dataItem)
        {
            try
            {
                SPListItem item = dataItem as SPListItem;
                if (item == null || string.IsNullOrEmpty(ImageField))
                    return "/_layouts/15/images/placeholder.png";
                
                try
                {
                    if (item.Fields.ContainsField(ImageField) && item[ImageField] != null)
                    {
                        // Get the field value as string
                        string fieldValue = item[ImageField].ToString();
                        
                        // If it's a URL, return it directly
                        if (fieldValue.StartsWith("http") || fieldValue.StartsWith("/"))
                        {
                            return fieldValue;
                        }
                        
                        // Try to extract URL from JSON or other formats
                        if (fieldValue.Contains("http") || fieldValue.Contains("/"))
                        {
                            // Extract URL using regex
                            Match match = Regex.Match(fieldValue, @"https?://[^\s"",<>]+");
                            if (match.Success)
                            {
                                return match.Value;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error getting image URL: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting image URL: {ex.Message}");
            }

            // Default image if no image found or on error
            return "/_layouts/15/images/placeholder.png";
        }

        // Simple array of month names
        private static readonly string[] _monthNames = {
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        };
        
        public string FormatDate(object dataItem)
        {
            try
            {
                // Get the item
                SPListItem item = dataItem as SPListItem;
                if (item == null || string.IsNullOrEmpty(DateField) || item[DateField] == null)
                    return string.Empty;
                
                // Get the date value
                object fieldValue = item[DateField];
                DateTime dateValue;
                
                // Handle different date types
                if (fieldValue is DateTime dt)
                {
                    dateValue = dt;
                }
                else if (!DateTime.TryParse(fieldValue.ToString(), out dateValue))
                {
                    // If we can't parse it, return empty
                    return string.Empty;
                }
                
                // Format with fixed month names
                int day = dateValue.Day;
                int month = dateValue.Month;
                int year = dateValue.Year;
                
                // Validate month index
                if (month < 1 || month > 12)
                    return string.Empty;
                
                // Return formatted date
                return day + " " + _monthNames[month - 1] + " " + year;
            }
            catch
            {
                // Silently handle any errors
                return string.Empty;
            }
        }

        // Dictionary to cache truncated descriptions for better performance
        private static Dictionary<string, string> _descriptionCache = new Dictionary<string, string>();
        
        public string GetDescription(object dataItem)
        {
            try
            {
                SPListItem item = dataItem as SPListItem;
                if (item == null || string.IsNullOrEmpty(DescriptionField) || item[DescriptionField] == null)
                    return string.Empty;
                
                // Create a cache key for this description
                string cacheKey = $"{item.ID}_{DescriptionField}";
                
                // Check if we have this description cached
                if (_descriptionCache.TryGetValue(cacheKey, out string cachedDesc))
                    return cachedDesc;
                
                // Get and truncate the description
                string description = TruncateDescription(item[DescriptionField].ToString());
                
                // Cache the result
                _descriptionCache[cacheKey] = description;
                
                return description;
            }
            catch
            {
                return string.Empty;
            }
        }
        
        public string TruncateDescription(string descText)
        {
            if (string.IsNullOrEmpty(descText))
                return string.Empty;

            // Limit description to 150 characters
            const int maxLength = 150;
            if (descText.Length <= maxLength)
                return descText;

            // Truncate at the last space before maxLength to avoid cutting words
            int lastSpace = descText.Substring(0, maxLength).LastIndexOf(' ');
            if (lastSpace <= 0) // No space found, just cut at maxLength
                lastSpace = maxLength;

            return descText.Substring(0, lastSpace) + "...";
        }

        // Cache for hyperlink information
        private static Dictionary<string, HyperlinkInfo> _hyperlinkCache = new Dictionary<string, HyperlinkInfo>();
        
        // Method to get the hyperlink URL and text from a SharePoint hyperlink field
        public HyperlinkInfo GetHyperlinkInfo(object dataItem)
        {
            try
            {
                SPListItem item = dataItem as SPListItem;
                if (item == null || string.IsNullOrEmpty(HyperlinkField))
                    return new HyperlinkInfo { Url = "#", DisplayText = "READ MORE..." };
                
                // Create a cache key for this hyperlink
                string cacheKey = $"{item.ID}_{HyperlinkField}";
                
                // Check if we have this hyperlink cached
                if (_hyperlinkCache.TryGetValue(cacheKey, out HyperlinkInfo cachedLink))
                    return cachedLink;
                
                // Default hyperlink info
                HyperlinkInfo result = new HyperlinkInfo { Url = "#", DisplayText = "READ MORE..." };
                
                // Get the field value
                if (item[HyperlinkField] != null)
                {
                    string fieldValue = item[HyperlinkField].ToString();
                    
                    // Quick parse for common formats
                    if (fieldValue.Contains(","))
                    {
                        string[] parts = fieldValue.Split(new char[] { ',' }, 2);
                        if (parts.Length == 2)
                        {
                            result = new HyperlinkInfo { Url = parts[0], DisplayText = parts[1] };
                        }
                    }
                    else if (fieldValue.StartsWith("http") || fieldValue.StartsWith("/"))
                    {
                        result = new HyperlinkInfo { Url = fieldValue, DisplayText = "READ MORE..." };
                    }
                }
                
                // Cache the result
                _hyperlinkCache[cacheKey] = result;
                return result;
            }
            catch
            {
                return new HyperlinkInfo { Url = "#", DisplayText = "READ MORE..." };
            }
        }
    }
    
    // Class to hold hyperlink information
    public class HyperlinkInfo
    {
        public string Url { get; set; }
        public string DisplayText { get; set; }
    }
}
