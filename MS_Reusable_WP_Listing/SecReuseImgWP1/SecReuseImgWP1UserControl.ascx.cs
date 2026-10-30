using System;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Data;
using Microsoft.SharePoint;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Utilities;

namespace MS_Reusable_WP_Listing.SecReuseImgWP1
{
    public partial class SecReuseImgWP1UserControl : UserControl
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

        // Property for the link field name (new property)
        private string _linkField = "Link";
        public string LinkField
        {
            get { return _linkField; }
            set { _linkField = value; }
        }

        // Property for the page size
        private int _pageSize = 6;
        public int PageSize
        {
            get { return _pageSize; }
            set { _pageSize = value; }
        }

        #endregion

        #region Page Events

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                // Set the page size for the data pager
                DataPager pager = lsvImages.FindControl("_dpPager") as DataPager;
                if (pager != null)
                {
                    pager.PageSize = PageSize;
                }

                // Load items from the SharePoint list
                LoadItems();
            }
        }

        #endregion

        #region Event Handlers

        protected void lsvImages_PagePropertiesChanging(object sender, PagePropertiesChangingEventArgs e)
        {
            // Set the start row index based on the page index
            DataPager pager = lsvImages.FindControl("_dpPager") as DataPager;
            if (pager != null)
            {
                pager.SetPageProperties(e.StartRowIndex, e.MaximumRows, false);
            }

            // Rebind the data
            LoadItems();
        }

        #endregion

        #region Helper Methods

        private void LoadItems()
        {
            try
            {
                if (string.IsNullOrEmpty(ListName))
                {
                    ShowError("Please configure the list name in the web part properties.");
                    return;
                }

                System.Diagnostics.Debug.WriteLine($"Loading items from list: {ListName}");
                System.Diagnostics.Debug.WriteLine($"Using fields - Title: {TitleField}, Description: {DescriptionField}, Image: {ImageField}, Date: {DateField}, Link: {LinkField}");

                // Set a timeout to prevent infinite loading
                int timeoutMs = 10000; // 10 seconds
                DateTime startTime = DateTime.Now;
                
                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                {
                    // Get the web based on the WebUrl property
                    SPWeb web = null;
                    
                    try
                    {
                        if (!string.IsNullOrEmpty(WebUrl))
                        {
                            // Use the specified web URL
                            web = site.OpenWeb(WebUrl);
                            System.Diagnostics.Debug.WriteLine($"Using web at URL: {WebUrl}");
                        }
                        else
                        {
                            // Use the current web
                            web = site.OpenWeb();
                            System.Diagnostics.Debug.WriteLine($"Using current web: {web.Url}");
                        }
                    }
                    catch (Exception webEx)
                    {
                        ShowError($"Error opening web: {webEx.Message}. Please check the Web URL value.");
                        return;
                    }
                    
                    if (web == null)
                    {
                        ShowError("Could not open the specified web. Please check the Web URL value.");
                        return;
                    }
                    
                    using (web)
                    {
                        // Check if the list exists
                        SPList list = web.Lists.TryGetList(ListName);
                        if (list == null)
                        {
                            // Check for timeout
                            if ((DateTime.Now - startTime).TotalMilliseconds > timeoutMs)
                            {
                                ShowError("Operation timed out. Please try again or check your configuration.");
                                return;
                            }
                            
                            // List all available lists for debugging
                            System.Text.StringBuilder availableLists = new System.Text.StringBuilder("Available lists: ");
                            foreach (SPList availableList in web.Lists)
                            {
                                availableLists.Append(availableList.Title + ", ");
                            }
                            System.Diagnostics.Debug.WriteLine(availableLists.ToString());

                            ShowError($"The list '{ListName}' does not exist. Please check the list name.");
                            return;
                        }

                        // Log the fields in the list
                        System.Text.StringBuilder listFields = new System.Text.StringBuilder("Fields in list: ");
                        foreach (SPField field in list.Fields)
                        {
                            listFields.Append(field.InternalName + ", ");
                        }
                        System.Diagnostics.Debug.WriteLine(listFields.ToString());

                        // Create a query to get the items
                        SPQuery query = new SPQuery();
                        
                        // Build the view fields - always include ID
                        string viewFields = "<FieldRef Name='ID' />";
                    
                        // Add title field
                        if (list.Fields.ContainsField(TitleField))
                        {
                            viewFields += $"<FieldRef Name='{TitleField}' />";
                        }
                        else
                        {
                            // Fallback to default Title field
                            viewFields += "<FieldRef Name='Title' />";
                        }
                    
                        // Add description field if specified
                        if (!string.IsNullOrEmpty(DescriptionField) && list.Fields.ContainsField(DescriptionField))
                        {
                            viewFields += $"<FieldRef Name='{DescriptionField}' />";
                        }
                    
                        // Add image field if specified
                        if (!string.IsNullOrEmpty(ImageField) && list.Fields.ContainsField(ImageField))
                        {
                            viewFields += $"<FieldRef Name='{ImageField}' />";
                        }
                        
                        // Always include the PublishingPageImage field if it exists
                        if (list.Fields.ContainsField("PublishingPageImage"))
                        {
                            viewFields += "<FieldRef Name='PublishingPageImage' />";
                        }
                        
                        // Also include Picture field as fallback
                        if (list.Fields.ContainsField("Picture"))
                        {
                            viewFields += "<FieldRef Name='Picture' />";
                        }
                    
                        // Add date field if specified
                        if (!string.IsNullOrEmpty(DateField) && list.Fields.ContainsField(DateField))
                        {
                            viewFields += $"<FieldRef Name='{DateField}' />";
                        }
                        
                        // Add link field if specified (new code)
                        if (!string.IsNullOrEmpty(LinkField) && list.Fields.ContainsField(LinkField))
                        {
                            viewFields += $"<FieldRef Name='{LinkField}' />";
                        }
                    
                        // Add document library fields
                        string[] docLibFields = new string[] { "FileRef", "FileDirRef", "FileLeafRef", "EncodedAbsUrl" };
                        foreach (string fieldName in docLibFields)
                        {
                            if (list.Fields.ContainsField(fieldName))
                            {
                                viewFields += $"<FieldRef Name='{fieldName}' />";
                            }
                        }
                    
                        query.ViewFields = viewFields;
                        System.Diagnostics.Debug.WriteLine($"View fields: {viewFields}");
                        
                        // Don't limit to only the view fields to ensure we get all needed fields
                        query.ViewFieldsOnly = false;
                        
                        // Order by the date field if specified, otherwise by ID
                        if (!string.IsNullOrEmpty(DateField) && list.Fields.ContainsField(DateField))
                        {
                            query.Query = $"<OrderBy><FieldRef Name='{DateField}' Ascending='False' /></OrderBy>";
                        }
                        else
                        {
                            query.Query = "<OrderBy><FieldRef Name='ID' Ascending='False' /></OrderBy>";
                        }

                        query.Query = ApplyApprovedItemsFilter(list, query.Query);
                    
                        // Check for timeout before getting items
                        if ((DateTime.Now - startTime).TotalMilliseconds > timeoutMs)
                        {
                            ShowError("Operation timed out while preparing query. Please try again or check your configuration.");
                            return;
                        }
                        
                        // Get the items
                        SPListItemCollection items = list.GetItems(query);
                        System.Diagnostics.Debug.WriteLine($"Found {items.Count} items in the list");
                        
                        // Check if there are any items
                        if (items.Count == 0)
                        {
                            lsvImages.EmptyDataTemplate = new EmptyDataTemplateHelper("No items found in the list.");
                            lsvImages.DataSource = null;
                            lsvImages.DataBind();
                            return;
                        }

                        // Debug the first item's fields
                        if (items.Count > 0)
                        {
                            // Check for timeout before processing items
                            if ((DateTime.Now - startTime).TotalMilliseconds > timeoutMs)
                            {
                                ShowError("Operation timed out while processing items. Please try again or check your configuration.");
                                return;
                            }
                            
                            SPListItem firstItem = items[0];
                            System.Text.StringBuilder firstItemFields = new System.Text.StringBuilder("First item fields: ");
                            foreach (SPField field in firstItem.Fields)
                            {
                                try
                                {
                                    string fieldName = field.InternalName;
                                    if (firstItem[fieldName] != null)
                                    {
                                        firstItemFields.Append($"{fieldName}='{firstItem[fieldName]}', ");
                                    }
                                }
                                catch { }
                            }
                            System.Diagnostics.Debug.WriteLine(firstItemFields.ToString());
                        }
                    
                        // Check for final timeout before binding
                        if ((DateTime.Now - startTime).TotalMilliseconds > timeoutMs)
                        {
                            ShowError("Operation timed out before data binding. Please try again or check your configuration.");
                            return;
                        }
                        
                        // Bind the items to the ListView
                        lsvImages.DataSource = items;
                        lsvImages.DataBind();
                        
                        // Show the content panel
                        pnlContent.Visible = true;
                        pnlError.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading items: {ex.Message}\n{ex.StackTrace}");
                ShowError($"Error loading items: {ex.Message}");
            }
        }

        private void ShowError(string message)
        {
            // Log the error for debugging
            System.Diagnostics.Debug.WriteLine($"ERROR: {message}");
            
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

        public string GetImageUrl(object dataItem)
        {
            try
            {
                SPListItem item = dataItem as SPListItem;
                if (item == null)
                    return "/_layouts/15/images/placeholder.png";
                
                // Try FileRef field first (common in document libraries)
                try
                {
                    if (item.Fields.ContainsField("FileRef") && item["FileRef"] != null)
                    {
                        string fileRef = item["FileRef"].ToString();
                        System.Diagnostics.Debug.WriteLine($"FileRef value: {fileRef}");
                        return fileRef;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error getting FileRef: {ex.Message}");
                }
                
                // Try EncodedAbsUrl field
                try
                {
                    if (item.Fields.ContainsField("EncodedAbsUrl") && item["EncodedAbsUrl"] != null)
                    {
                        string url = item["EncodedAbsUrl"].ToString();
                        System.Diagnostics.Debug.WriteLine($"EncodedAbsUrl value: {url}");
                        return url;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error getting EncodedAbsUrl: {ex.Message}");
                }
                
                // Try FileLeafRef field (filename) with FileDirRef (directory)
                try
                {
                    if (item.Fields.ContainsField("FileLeafRef") && item["FileLeafRef"] != null)
                    {
                        string fileName = item["FileLeafRef"].ToString();
                        string folderUrl = item.Web.Url;
                        if (item.Fields.ContainsField("FileDirRef") && item["FileDirRef"] != null)
                        {
                            folderUrl = item["FileDirRef"].ToString();
                        }
                        string fileUrl = folderUrl + "/" + fileName;
                        System.Diagnostics.Debug.WriteLine($"Constructed URL from FileLeafRef: {fileUrl}");
                        return fileUrl;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error getting FileLeafRef: {ex.Message}");
                }
                
                // Fall back to the configured ImageField
                if (string.IsNullOrEmpty(ImageField))
                    return "/_layouts/15/images/placeholder.png";

                try
                {
                    // Check if the image field exists and has a value
                    if (item.Fields.ContainsField(ImageField) && item[ImageField] != null)
                    {
                        string fieldValue = item[ImageField].ToString();
                        
                        // Debug the field value
                        System.Diagnostics.Debug.WriteLine($"Image field value for {item.ID}: {fieldValue}");
                        
                        // Handle different types of image fields
                        if (string.IsNullOrEmpty(fieldValue))
                        {
                            return "/_layouts/15/images/placeholder.png";
                        }
                        
                        // Check if it's a URL field
                        if (item.Fields[ImageField].Type == SPFieldType.URL)
                        {
                            try
                            {
                                SPFieldUrlValue urlValue = new SPFieldUrlValue(fieldValue);
                                return urlValue.Url;
                            }
                            catch
                            {
                                // If parsing as URL fails, try other methods
                            }
                        }
                        
                        // Check if it's a Picture field (contains URL:)
                        if (fieldValue.Contains("URL:"))
                        {
                            try
                            {
                                int startIndex = fieldValue.IndexOf("URL:") + 4;
                                int endIndex = fieldValue.IndexOf(",", startIndex);
                                if (endIndex == -1) endIndex = fieldValue.Length;
                                return fieldValue.Substring(startIndex, endIndex - startIndex).Trim();
                            }
                            catch
                            {
                                // If extraction fails, continue to next method
                            }
                        }
                        
                        // Check if it's a JSON string containing URL
                        if (fieldValue.Contains("\"Url\":") || fieldValue.Contains("\"url\":"))
                        {
                            try
                            {
                                // Try to extract URL from JSON-like string
                                int urlIndex = fieldValue.IndexOf("\"Url\":");
                                if (urlIndex == -1) urlIndex = fieldValue.IndexOf("\"url\":");
                                
                                if (urlIndex >= 0)
                                {
                                    int valueStart = fieldValue.IndexOf('"', urlIndex + 6) + 1;
                                    int valueEnd = fieldValue.IndexOf('"', valueStart);
                                    if (valueEnd > valueStart)
                                    {
                                        return fieldValue.Substring(valueStart, valueEnd - valueStart);
                                    }
                                }
                            }
                            catch
                            {
                                // If JSON parsing fails, continue to next method
                            }
                        }
                        
                        // If it looks like a URL, return it directly
                        if (fieldValue.StartsWith("http") || fieldValue.StartsWith("/"))
                        {
                            return fieldValue;
                        }
                        
                        // Last resort: try to find any URL-like pattern in the string
                        if (fieldValue.Contains("http") || fieldValue.Contains("/"))
                        {
                            try
                            {
                                int httpIndex = fieldValue.IndexOf("http");
                                if (httpIndex >= 0)
                                {
                                    int endIndex = fieldValue.IndexOf('"', httpIndex);
                                    if (endIndex == -1) endIndex = fieldValue.IndexOf(' ', httpIndex);
                                    if (endIndex == -1) endIndex = fieldValue.Length;
                                    return fieldValue.Substring(httpIndex, endIndex - httpIndex);
                                }
                                
                                int slashIndex = fieldValue.IndexOf("/");
                                if (slashIndex >= 0 && slashIndex < 10) // Only consider if slash is near the beginning
                                {
                                    return fieldValue;
                                }
                            }
                            catch
                            {
                                // If URL extraction fails, use default
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error getting image URL: {ex.Message}");
                }

                // Default image if no image found or on error
                return "/_layouts/15/images/placeholder.png";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in GetImageUrl: {ex.Message}");
                return "/_layouts/15/images/placeholder.png";
            }
        }

        public string FormatDate(object dataItem)
        {
            SPListItem item = dataItem as SPListItem;
            if (item == null || string.IsNullOrEmpty(DateField))
                return string.Empty;

            try
            {
                // Check if the date field exists and has a value
                if (item.Fields.ContainsField(DateField) && item[DateField] != null)
                {
                    // Try to parse the date value
                    if (DateTime.TryParse(item[DateField].ToString(), out DateTime dateValue))
                    {
                        return dateValue.ToString("dd MMMM yyyy");
                    }
                }
            }
            catch
            {
                // Return empty string on error
                return string.Empty;
            }

            return string.Empty;
        }

        public string GetDescription(object dataItem)
        {
            try
            {
                SPListItem item = dataItem as SPListItem;
                if (item == null)
                    return string.Empty;
                
                // Try to get description from the configured field
                if (!string.IsNullOrEmpty(DescriptionField) && item.Fields.ContainsField(DescriptionField) && item[DescriptionField] != null)
                {
                    return TruncateDescription(item[DescriptionField]);
                }
                
                // Try common description field names
                string[] commonDescriptionFields = new string[] { "Description", "Body", "Content", "Text", "Comments", "Details" };
                foreach (string fieldName in commonDescriptionFields)
                {
                    if (item.Fields.ContainsField(fieldName) && item[fieldName] != null)
                    {
                        return TruncateDescription(item[fieldName]);
                    }
                }
                
                // If no description found, use title or empty string
                if (item.Fields.ContainsField("Title") && item["Title"] != null)
                {
                    return TruncateDescription(item["Title"]);
                }
                
                return string.Empty;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting description: {ex.Message}");
                return string.Empty;
            }
        }
        
        public string TruncateDescription(object description)
        {
            try
            {
                if (description == null)
                    return string.Empty;

                string descText = description.ToString();
                if (string.IsNullOrEmpty(descText))
                    return string.Empty;

                // Debug the description value
                System.Diagnostics.Debug.WriteLine($"Description before truncation: {descText.Substring(0, Math.Min(50, descText.Length))}...");

                // Limit description to 150 characters
                int maxLength = 150;
                if (descText.Length <= maxLength)
                    return descText;

                // Truncate at the last space before maxLength to avoid cutting words
                int lastSpace = descText.Substring(0, maxLength).LastIndexOf(' ');
                if (lastSpace <= 0) // No space found, just cut at maxLength
                    lastSpace = maxLength;

                return descText.Substring(0, lastSpace) + "...";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error truncating description: {ex.Message}");
                return string.Empty;
            }
        }

        // New method to get the link URL from the specified field
        public string GetLinkUrl(object dataItem)
        {
            try
            {
                SPListItem item = dataItem as SPListItem;
                if (item == null || string.IsNullOrEmpty(LinkField))
                    return "#"; // Default fallback URL
                
                // Check if the link field exists and has a value
                if (item.Fields.ContainsField(LinkField) && item[LinkField] != null)
                {
                    string fieldValue = item[LinkField].ToString();
                    
                    // Debug the field value
                    System.Diagnostics.Debug.WriteLine($"Link field value for {item.ID}: {fieldValue}");
                    
                    // If the field is empty, return default
                    if (string.IsNullOrEmpty(fieldValue))
                    {
                        return "#";
                    }
                    
                    // Check if it's a URL field
                    if (item.Fields[LinkField].Type == SPFieldType.URL)
                    {
                        try
                        {
                            SPFieldUrlValue urlValue = new SPFieldUrlValue(fieldValue);
                            return urlValue.Url;
                        }
                        catch
                        {
                            // If parsing as URL fails, try other methods
                        }
                    }
                    
                    // If it looks like a URL, return it directly
                    if (fieldValue.StartsWith("http") || fieldValue.StartsWith("/"))
                    {
                        return fieldValue;
                    }
                    
                    // Check if it's a format like "URL|Description" or "URL,Description"
                    char[] separators = new char[] { '|', ',' };
                    foreach (char separator in separators)
                    {
                        if (fieldValue.Contains(separator))
                        {
                            string[] parts = fieldValue.Split(separator);
                            if (parts.Length > 0 && !string.IsNullOrEmpty(parts[0]))
                            {
                                return parts[0].Trim();
                            }
                        }
                    }
                    
                    // Last resort: try to find any URL-like pattern in the string
                    if (fieldValue.Contains("http"))
                    {
                        try
                        {
                            int httpIndex = fieldValue.IndexOf("http");
                            if (httpIndex >= 0)
                            {
                                int endIndex = fieldValue.IndexOf('"', httpIndex);
                                if (endIndex == -1) endIndex = fieldValue.IndexOf(' ', httpIndex);
                                if (endIndex == -1) endIndex = fieldValue.Length;
                                return fieldValue.Substring(httpIndex, endIndex - httpIndex);
                            }
                        }
                        catch
                        {
                            // If URL extraction fails, use default
                        }
                    }
                    
                    // If all else fails, just return the field value as is
                    return fieldValue;
                }
                
                // If the field doesn't exist or has no value, return default
                return "#";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting link URL: {ex.Message}");
                return "#";
            }
        }

        #endregion
    }

    // Helper class for empty data template
    public class EmptyDataTemplateHelper : ITemplate
    {
        private string _message;

        public EmptyDataTemplateHelper(string message)
        {
            _message = message;
        }

        public void InstantiateIn(Control container)
        {
            Literal literal = new Literal();
            literal.Text = $"<div class='no-data'>{_message}</div>";
            container.Controls.Add(literal);
        }
    }
}
