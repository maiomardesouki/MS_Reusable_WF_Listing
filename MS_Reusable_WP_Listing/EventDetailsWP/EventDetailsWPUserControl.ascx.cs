using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using System.Globalization;

namespace MS_Reusable_WP_Listing.EventDetailsWP
{
    public partial class EventDetailsWPUserControl : UserControl
    {
        #region Properties

        // Property to store the list name from the web part properties
        private string _listName = string.Empty;
        public string ListName
        {
            get { return _listName; }
            set { _listName = value; }
        }
        
        // Property to store the title field name
        private string _titleField = "Title";
        public string TitleField
        {
            get { return _titleField; }
            set { _titleField = value; }
        }
        
        // Property to store the Arabic title field name
        private string _titleArField = "TitleAr";
        public string TitleArField
        {
            get { return _titleArField; }
            set { _titleArField = value; }
        }
        
        // Property to store the description field name
        private string _descriptionField = "Description";
        public string DescriptionField
        {
            get { return _descriptionField; }
            set { _descriptionField = value; }
        }
        
        // Property to store the Arabic description field name
        private string _descriptionArField = "DescriptionAr";
        public string DescriptionArField
        {
            get { return _descriptionArField; }
            set { _descriptionArField = value; }
        }
        
        // Property to store the location field name
        private string _locationField = "Location";
        public string LocationField
        {
            get { return _locationField; }
            set { _locationField = value; }
        }
        
        // Property to store the Arabic location field name
        private string _locationArField = "LocationAr";
        public string LocationArField
        {
            get { return _locationArField; }
            set { _locationArField = value; }
        }
        
        // Property to store the start date field name
        private string _startDateField = "StartDate";
        public string StartDateField
        {
            get { return _startDateField; }
            set { _startDateField = value; }
        }
        
        // Property to store the end date field name
        private string _endDateField = "EndDate";
        public string EndDateField
        {
            get { return _endDateField; }
            set { _endDateField = value; }
        }
        
        // Property to store the category field name
        private string _categoryField = "CalendarCategory";
        public string CategoryField
        {
            get { return _categoryField; }
            set { _categoryField = value; }
        }
        
        // Property to store the approval field name
        private string _approvedField = "IsApproved";
        public string ApprovedField
        {
            get { return _approvedField; }
            set { _approvedField = value; }
        }
        
        // Property to determine if we should use Arabic fields
        private bool _useArabicFields = false;
        public bool UseArabicFields
        {
            get { return _useArabicFields; }
            set { _useArabicFields = value; }
        }

        #endregion

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                // Check if we have a list name in the query string and override the property if so
                string listNameParam = Request.QueryString["ListName"];
                if (!string.IsNullOrEmpty(listNameParam))
                {
                    ListName = listNameParam;
                }
                
                // Get the event ID from the query string
                string eventIdParam = Request.QueryString["EventID"];
                int eventId = 0;

                if (!string.IsNullOrEmpty(eventIdParam) && int.TryParse(eventIdParam, out eventId))
                {
                    // Load the event details
                    LoadEventDetails(eventId);
                }
                else
                {
                    // No event ID provided, show error message
                    DisplayErrorMessage("No event information provided.");
                }
                
                // Check if we should use Arabic fields based on query string
                string useArabicParam = Request.QueryString["UseArabic"];
                if (!string.IsNullOrEmpty(useArabicParam))
                {
                    bool useArabic = false;
                    if (bool.TryParse(useArabicParam, out useArabic))
                    {
                        UseArabicFields = useArabic;
                    }
                }
            }
        }

        private void LoadEventDetails(int eventId)
        {
            try
            {
                // Determine which list to use
                string listToUse = !string.IsNullOrEmpty(ListName) ? ListName : "CalendarEvents";

                string listWebUrlParam = Request.QueryString["ListWebUrl"];

                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                {
                    using (SPWeb currentWeb = site.OpenWeb(SPContext.Current.Web.ID))
                    {
                        if (TrySetEventDetailsFromWeb(currentWeb, listToUse, eventId))
                            return;
                    }

                    using (SPWeb rootWeb = site.RootWeb)
                    {
                        if (TrySetEventDetailsFromWeb(rootWeb, listToUse, eventId))
                            return;
                    }
                }

                if (!string.IsNullOrEmpty(listWebUrlParam))
                {
                    try
                    {
                        using (SPSite listSite = new SPSite(listWebUrlParam))
                        using (SPWeb listWeb = listSite.OpenWeb())
                        {
                            if (TrySetEventDetailsFromWeb(listWeb, listToUse, eventId))
                                return;
                        }
                    }
                    catch
                    {
                    }
                }

                try
                {
                    string authority = Request.Url.GetLeftPart(UriPartial.Authority);
                    using (SPSite rootSite = new SPSite(authority))
                    using (SPWeb rootWebAppWeb = rootSite.RootWeb)
                    {
                        if (TrySetEventDetailsFromWeb(rootWebAppWeb, listToUse, eventId))
                            return;
                    }
                }
                catch
                {
                }

                DisplayErrorMessage("List not found.");
            }
            catch (Exception ex)
            {
                DisplayErrorMessage("Error loading event details.");
            }
        }

        private bool TrySetEventDetailsFromWeb(SPWeb web, string listNameOrUrl, int eventId)
        {
            SPList list = ResolveList(web, listNameOrUrl);
            if (list == null)
                return false;

            SPListItem item;
            try
            {
                item = list.GetItemById(eventId);
            }
            catch
            {
                item = null;
            }

            if (item == null)
                return false;

            if (list.EnableModeration)
            {
                try
                {
                    if (item.ModerationInformation != null && item.ModerationInformation.Status != SPModerationStatusType.Approved)
                        return false;
                }
                catch
                {
                    return false;
                }
            }

            SetEventDetails(item);
            return true;
        }

        private static SPList ResolveList(SPWeb web, string listNameOrUrl)
        {
            if (web == null)
                return null;

            if (string.IsNullOrWhiteSpace(listNameOrUrl))
                return null;

            SPList list = web.Lists.TryGetList(listNameOrUrl);
            if (list != null)
                return list;

            string key = listNameOrUrl.Trim();

            if (Uri.TryCreate(key, UriKind.Absolute, out Uri uri))
                key = uri.AbsolutePath;

            key = key.Trim().TrimEnd('/');
            string[] segments = key.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

            if (segments.Length > 0)
            {
                string last = segments[segments.Length - 1];
                if (last.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase) && segments.Length >= 2)
                    key = segments[segments.Length - 2];
                else
                    key = last;
            }

            string normalizedKey = NormalizeListKey(key);

            foreach (SPList candidate in web.Lists)
            {
                try
                {
                    if (NormalizeListKey(candidate.Title) == normalizedKey)
                        return candidate;

                    if (NormalizeListKey(candidate.RootFolder != null ? candidate.RootFolder.Name : string.Empty) == normalizedKey)
                        return candidate;

                    if (candidate.RootFolder != null && NormalizeListKey(candidate.RootFolder.Url) == normalizedKey)
                        return candidate;

                    if (candidate.RootFolder != null && candidate.RootFolder.Url != null && candidate.RootFolder.Url.IndexOf("/Lists/" + key, StringComparison.OrdinalIgnoreCase) >= 0)
                        return candidate;
                }
                catch
                {
                }
            }

            return null;
        }

        private static string NormalizeListKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            string trimmed = value.Trim().TrimEnd('/');
            string[] segments = trimmed.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string last = segments.Length > 0 ? segments[segments.Length - 1] : trimmed;
            return new string(last.Where(c => !char.IsWhiteSpace(c)).ToArray()).ToLowerInvariant();
        }

        private void SetEventDetails(SPListItem item)
        {
            try
            {
                // Set the title based on language preference
                string titleValue = "[No Title]";
                
                // Try to get the title in the appropriate language
                if (UseArabicFields && item.Fields.ContainsField(TitleArField) && item[TitleArField] != null)
                {
                    titleValue = item[TitleArField].ToString();
                }
                else if (item.Fields.ContainsField(TitleField) && item[TitleField] != null)
                {
                    titleValue = item[TitleField].ToString();
                }
                
                litTitle.Text = titleValue;

                // Set the date
                if (item.Fields.ContainsField(StartDateField) && item[StartDateField] != null && 
                    DateTime.TryParse(item[StartDateField].ToString(), out DateTime startDate))
                {
                    if (item.Fields.ContainsField(EndDateField) && item[EndDateField] != null && 
                        DateTime.TryParse(item[EndDateField].ToString(), out DateTime endDate))
                    {
                        // Format the date range
                        if (startDate.Date == endDate.Date)
                        {
                            // Same day event
                            litDate.Text = startDate.ToString("MMMM d, yyyy", CultureInfo.InvariantCulture) + " | " + 
                                          startDate.ToString("h:mm tt", CultureInfo.InvariantCulture) + " - " + 
                                          endDate.ToString("h:mm tt", CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            // Multi-day event
                            litDate.Text = startDate.ToString("MMMM d, yyyy h:mm tt", CultureInfo.InvariantCulture) + " - " + 
                                          endDate.ToString("MMMM d, yyyy h:mm tt", CultureInfo.InvariantCulture);
                        }
                    }
                    else
                    {
                        // Only start date available
                        litDate.Text = startDate.ToString("MMMM d, yyyy h:mm tt", CultureInfo.InvariantCulture);
                    }
                }
                else
                {
                    litDate.Text = "[No Date]";
                }

                // Set the description based on language preference
                string description = GetDescriptionFromItem(item);
                litDescription.Text = description;
                
                // Set the location if available
                string locationValue = string.Empty;
                string locationFieldToUse = UseArabicFields && item.Fields.ContainsField(LocationArField) ? 
                                           LocationArField : LocationField;
                
                if (item.Fields.ContainsField(locationFieldToUse) && item[locationFieldToUse] != null)
                {
                    locationValue = item[locationFieldToUse].ToString();
                    // Add location to the date display if available
                    if (!string.IsNullOrEmpty(locationValue))
                    {
                        litDate.Text += "<br/><span class=\"event-location\"><i class=\"fas fa-map-marker-alt\"></i> " + 
                                       locationValue + "</span>";
                    }
                }

                // Category is no longer displayed in the UI
                litCategory.Text = string.Empty;
                
                // Check if the event is approved (if applicable)
                bool isApproved = true; // Default to true if field doesn't exist
                if (item.Fields.ContainsField(ApprovedField) && item[ApprovedField] != null)
                {
                    try
                    {
                        isApproved = Convert.ToBoolean(item[ApprovedField]);
                    }
                    catch { /* Use default value if conversion fails */ }
                }
                
                // If not approved, we can indicate this in a different way without adding text to the title
                if (!isApproved)
                {
                    // Add the pending approval tag to the date section instead
                    litDate.Text += "<span class=\"pending-tag\">Pending Approval</span>";
                }

                // Set external link if available
                if (item.Fields.ContainsField("ExternalLink") && item["ExternalLink"] != null && !string.IsNullOrEmpty(item["ExternalLink"].ToString()))
                {
                    string externalUrl = item["ExternalLink"].ToString();
                    litExternalLink.Text = $"<a href=\"{externalUrl}\" target=\"_blank\" class=\"btn-external\"><i class=\"fas fa-external-link-alt\"></i> External Link</a>";
                    divExternalLink.Visible = true;
                }
                else
                {
                    litExternalLink.Text = string.Empty;
                    divExternalLink.Visible = false;
                }

                // Handle attachments
                HandleAttachments(item);

                // Handle links
                HandleLinks(item);

                // Handle images
                HandleImages(item);
            }
            catch (Exception ex)
            {
                DisplayErrorMessage("Error setting event details.");
            }
        }

        private string GetDescriptionFromItem(SPListItem item)
        {
            // First try to get the description based on language preference
            if (UseArabicFields && item.Fields.ContainsField(DescriptionArField) && 
                item[DescriptionArField] != null && !string.IsNullOrEmpty(item[DescriptionArField].ToString()))
            {
                return item[DescriptionArField].ToString();
            }
            else if (item.Fields.ContainsField(DescriptionField) && 
                     item[DescriptionField] != null && !string.IsNullOrEmpty(item[DescriptionField].ToString()))
            {
                return item[DescriptionField].ToString();
            }
            
            // If primary description fields aren't available, try other possible fields
            string[] possibleDescriptionFields = new string[] { "Body", "Content", "Text", "Comments", "Details" };
            
            foreach (string fieldName in possibleDescriptionFields)
            {
                try
                {
                    if (item.Fields.ContainsField(fieldName) && item[fieldName] != null && !string.IsNullOrEmpty(item[fieldName].ToString()))
                    {
                        return item[fieldName].ToString();
                    }
                }
                catch { }
            }
            
            // If no description found, return the title or empty string
            string titleFieldToUse = UseArabicFields && item.Fields.ContainsField(TitleArField) ? TitleArField : TitleField;
            return item.Fields.ContainsField(titleFieldToUse) && item[titleFieldToUse] != null ? 
                   item[titleFieldToUse].ToString() : string.Empty;
        }

        private void HandleAttachments(SPListItem item)
        {
            try
            {
                if (item.Attachments.Count > 0)
                {
                    divAttachments.Visible = true;
                    string attachmentsHtml = string.Empty;

                    foreach (string attachment in item.Attachments)
                    {
                        string attachmentUrl = item.Web.Url + "/_layouts/15/download.aspx?SourceUrl=" + 
                                              SPEncode.UrlEncode(item.Web.Url + "/Lists/" + item.ParentList.Title + "/Attachments/" + 
                                              item.ID + "/" + attachment);
                        
                        string fileExtension = System.IO.Path.GetExtension(attachment).ToLower();
                        string iconClass = GetFileIconClass(fileExtension);
                        
                        attachmentsHtml += $"<li><a href=\"{attachmentUrl}\" target=\"_blank\"><i class=\"{iconClass}\"></i> {attachment}</a></li>";
                    }

                    litAttachments.Text = attachmentsHtml;
                }
                else
                {
                    divAttachments.Visible = false;
                }
            }
            catch
            {
                divAttachments.Visible = false;
            }
        }

        private void HandleLinks(SPListItem item)
        {
            try
            {
                // Check if the item has a Links field
                bool hasLinks = false;
                string linksHtml = string.Empty;

                // Try to get links from a dedicated field if it exists
                if (item.Fields.ContainsField("Links") && item["Links"] != null)
                {
                    string linkText = item["Links"].ToString();
                    if (!string.IsNullOrEmpty(linkText))
                    {
                        hasLinks = true;
                        
                        // Check if the link text contains multiple URLs (separated by line breaks or semicolons)
                        string[] possibleSeparators = new string[] { "\n", ";", "," };
                        bool hasSeparator = false;
                        
                        foreach (string separator in possibleSeparators)
                        {
                            if (linkText.Contains(separator))
                            {
                                hasSeparator = true;
                                string[] links = linkText.Split(new string[] { separator }, StringSplitOptions.RemoveEmptyEntries);
                                
                                foreach (string link in links)
                                {
                                    string trimmedLink = link.Trim();
                                    if (!string.IsNullOrEmpty(trimmedLink))
                                    {
                                        // Try to extract a title if the link has a format like "Title | URL"
                                        string title = trimmedLink;
                                        string url = trimmedLink;
                                        
                                        if (trimmedLink.Contains("|"))
                                        {
                                            string[] parts = trimmedLink.Split('|');
                                            if (parts.Length >= 2)
                                            {
                                                title = parts[0].Trim();
                                                url = parts[1].Trim();
                                            }
                                        }
                                        
                                        // Ensure URL has http/https prefix
                                        if (!url.StartsWith("http://") && !url.StartsWith("https://"))
                                        {
                                            url = "http://" + url;
                                        }
                                        
                                        linksHtml += $"<li><a href=\"{url}\" target=\"_blank\"><i class=\"fas fa-link\"></i> {title}</a></li>";
                                    }
                                }
                                
                                break; // Found a separator, no need to check others
                            }
                        }
                        
                        // If no separator found, treat as a single link
                        if (!hasSeparator)
                        {
                            string url = linkText;
                            
                            // Ensure URL has http/https prefix
                            if (!url.StartsWith("http://") && !url.StartsWith("https://"))
                            {
                                url = "http://" + url;
                            }
                            
                            linksHtml += $"<li><a href=\"{url}\" target=\"_blank\"><i class=\"fas fa-link\"></i> {linkText}</a></li>";
                        }
                    }
                }

                // Set the links HTML if we found any
                if (hasLinks)
                {
                    divLinks.Visible = true;
                    litLinks.Text = linksHtml;
                }
                else
                {
                    divLinks.Visible = false;
                }
            }
            catch
            {
                divLinks.Visible = false;
            }
        }

        private void HandleImages(SPListItem item)
        {
            try
            {
                string imagesHtml = string.Empty;
                bool hasImages = false;

                // Check for image field
                if (item.Fields.ContainsField("Image") && item["Image"] != null)
                {
                    string imageUrlStr = item["Image"].ToString();
                    if (!string.IsNullOrEmpty(imageUrlStr))
                    {
                        string[] imageEntries = imageUrlStr.Split(new[] { "\r\n", "\n", ";" }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string entry in imageEntries)
                        {
                            string trimmedEntry = entry.Trim();
                            if (string.IsNullOrEmpty(trimmedEntry))
                                continue;

                            string urlToRender = trimmedEntry;

                            try
                            {
                                SPFieldUrlValue urlValue = new SPFieldUrlValue(trimmedEntry);
                                if (!string.IsNullOrEmpty(urlValue.Url))
                                {
                                    urlToRender = urlValue.Url;
                                }
                            }
                            catch
                            {
                                int firstCommaIndex = trimmedEntry.IndexOf(',');
                                if (firstCommaIndex > 0)
                                {
                                    string possibleUrl = trimmedEntry.Substring(0, firstCommaIndex).Trim();
                                    if (!string.IsNullOrEmpty(possibleUrl))
                                    {
                                        urlToRender = possibleUrl;
                                    }
                                }
                            }

                            if (string.IsNullOrEmpty(urlToRender))
                                continue;

                            // Ensure URL has http/https prefix if it's not a site-relative URL
                            if (!urlToRender.StartsWith("/") && !urlToRender.StartsWith("http://") && !urlToRender.StartsWith("https://"))
                            {
                                urlToRender = "http://" + urlToRender;
                            }

                            // Guard against rendering description-only fragments as images
                            if (!(urlToRender.StartsWith("/") || urlToRender.StartsWith("http://") || urlToRender.StartsWith("https://")))
                                continue;

                            imagesHtml += $"<img src=\"{urlToRender}\" alt=\"{litTitle.Text}\" class=\"event-image\" />";
                            hasImages = true;
                        }
                    }
                }

                // Set the images HTML and visibility
                if (hasImages)
                {
                    litImages.Text = imagesHtml;
                    divImages.Visible = true;
                }
                else
                {
                    litImages.Text = string.Empty;
                    divImages.Visible = false;
                }
            }
            catch
            {
                litImages.Text = string.Empty;
                divImages.Visible = false;
            }
        }

        private string GetFileIconClass(string fileExtension)
        {
            // Map file extensions to Font Awesome icon classes
            switch (fileExtension)
            {
                case ".pdf":
                    return "far fa-file-pdf";
                case ".doc":
                case ".docx":
                    return "far fa-file-word";
                case ".xls":
                case ".xlsx":
                    return "far fa-file-excel";
                case ".ppt":
                case ".pptx":
                    return "far fa-file-powerpoint";
                case ".jpg":
                case ".jpeg":
                case ".png":
                case ".gif":
                    return "far fa-file-image";
                case ".zip":
                case ".rar":
                    return "far fa-file-archive";
                default:
                    return "far fa-file";
            }
        }

        private void DisplayErrorMessage(string message)
        {
            // Clear all content
            litTitle.Text = "Error";
            litDate.Text = string.Empty;
            litDescription.Text = $"<div class=\"error-message\">{message}</div>";
            litExternalLink.Text = string.Empty;
            litAttachments.Text = string.Empty;
            litLinks.Text = string.Empty;
            litImages.Text = string.Empty;
            
            // Hide all optional sections
            divAttachments.Visible = false;
            divLinks.Visible = false;
            divImages.Visible = false;
            divExternalLink.Visible = false;
        }
    }
}
