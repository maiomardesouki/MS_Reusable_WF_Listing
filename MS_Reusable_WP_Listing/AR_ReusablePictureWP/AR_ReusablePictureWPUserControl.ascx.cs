using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.SharePoint;

namespace MS_Reusable_WP_Listing.AR_ReusablePictureWP
{
    public partial class AR_ReusablePictureWPUserControl : UserControl
    {
        private string _listName = string.Empty;
        public string ListName
        {
            get { return _listName; }
            set { _listName = value; }
        }

        private string _webUrl = string.Empty;
        public string WebUrl
        {
            get { return _webUrl; }
            set { _webUrl = value; }
        }

        private string _titleField = "Title";
        public string TitleField
        {
            get { return _titleField; }
            set { _titleField = value; }
        }

        private string _descriptionField = "Description";
        public string DescriptionField
        {
            get { return _descriptionField; }
            set { _descriptionField = value; }
        }

        private string _imageField = "Picture";
        public string ImageField
        {
            get { return _imageField; }
            set { _imageField = value; }
        }

        private string _dateField = "Date";
        public string DateField
        {
            get { return _dateField; }
            set { _dateField = value; }
        }

        private string _hyperlinkField = "Link";
        public string HyperlinkField
        {
            get { return _hyperlinkField; }
            set { _hyperlinkField = value; }
        }

        private int _pageSize = 6;
        public int PageSize
        {
            get { return _pageSize; }
            set { _pageSize = value; }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                DataPager pager = lsvPictures.FindControl("_dpPager") as DataPager;
                if (pager != null)
                {
                    pager.PageSize = PageSize;
                }

                LoadItems();
            }
        }

        protected void lsvPictures_PagePropertiesChanging(object sender, PagePropertiesChangingEventArgs e)
        {
            DataPager pager = lsvPictures.FindControl("_dpPager") as DataPager;
            if (pager != null)
            {
                pager.SetPageProperties(e.StartRowIndex, e.MaximumRows, false);
                LoadItems();
            }
        }

        private static SPListItemCollection _cachedItems;
        private static DateTime _cacheExpiry = DateTime.MinValue;
        private static string _cacheKey = string.Empty;

        private void LoadItems()
        {
            try
            {
                if (string.IsNullOrEmpty(ListName))
                {
                    ShowError("يرجى تحديد اسم القائمة من خصائص الويب بارت.");
                    return;
                }

                string currentCacheKey = $"{ListName}_{WebUrl}_{TitleField}_{DescriptionField}_{DateField}_{HyperlinkField}";
                if (_cachedItems != null && DateTime.Now < _cacheExpiry && currentCacheKey == _cacheKey)
                {
                    lsvPictures.DataSource = _cachedItems;
                    lsvPictures.DataBind();
                    pnlContent.Visible = true;
                    pnlError.Visible = false;
                    return;
                }

                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                {
                    SPWeb web = string.IsNullOrEmpty(WebUrl) ? site.OpenWeb() : site.OpenWeb(WebUrl);
                    using (web)
                    {
                        SPList list = web.Lists.TryGetList(ListName);
                        if (list == null)
                        {
                            ShowError("القائمة '" + ListName + "' غير موجودة في الموقع '" + web.Url + "'.");
                            return;
                        }

                        string resolvedTitleField = ResolveFieldInternalName(list, TitleField);
                        string resolvedDescriptionField = ResolveFieldInternalName(list, DescriptionField);
                        string resolvedDateField = ResolveFieldInternalName(list, DateField);
                        string resolvedHyperlinkField = ResolveFieldInternalName(list, HyperlinkField);

                        SPQuery query = new SPQuery();
                        query.ViewFields = "<FieldRef Name='ID' /><FieldRef Name='Title' />";

                        if (!string.IsNullOrEmpty(resolvedTitleField) && !string.Equals(resolvedTitleField, "Title", StringComparison.OrdinalIgnoreCase))
                            query.ViewFields += "<FieldRef Name='" + resolvedTitleField + "' />";

                        if (!string.IsNullOrEmpty(resolvedDescriptionField))
                            query.ViewFields += "<FieldRef Name='" + resolvedDescriptionField + "' />";

                        if (!string.IsNullOrEmpty(resolvedDateField))
                        {
                            query.ViewFields += "<FieldRef Name='" + resolvedDateField + "' />";
                            query.Query = "<OrderBy><FieldRef Name='" + resolvedDateField + "' Ascending='False' /></OrderBy>";
                        }

                        if (!string.IsNullOrEmpty(resolvedHyperlinkField))
                            query.ViewFields += "<FieldRef Name='" + resolvedHyperlinkField + "' />";

                        query.ViewFieldsOnly = true;

                        query.Query = ApplyApprovedItemsFilter(list, query.Query);

                        SPListItemCollection items = list.GetItems(query);
                        if (items.Count == 0)
                        {
                            ShowError("لا توجد عناصر في القائمة.", false);
                            return;
                        }

                        _cachedItems = items;
                        _cacheExpiry = DateTime.Now.AddMinutes(5);
                        _cacheKey = currentCacheKey;

                        lsvPictures.DataSource = items;
                        lsvPictures.DataBind();

                        pnlContent.Visible = true;
                        pnlError.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("حدث خطأ أثناء تحميل العناصر: " + ex.Message);
            }
        }

        private void ShowError(string message, bool isError = true)
        {
            Debug.WriteLine($"{(isError ? "ERROR" : "INFO")}: {message}");
            litError.Text = message;
            pnlError.Visible = true;
            pnlContent.Visible = false;
            UpdatePanel1?.Update();
        }

        private static string ResolveFieldInternalName(SPList list, string fieldName)
        {
            if (list == null || string.IsNullOrWhiteSpace(fieldName))
                return null;

            if (list.Fields.ContainsField(fieldName))
                return fieldName;

            SPField match = list.Fields
                .Cast<SPField>()
                .FirstOrDefault(f =>
                    string.Equals(f.InternalName, fieldName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(f.StaticName, fieldName, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(f.Title, fieldName, StringComparison.OrdinalIgnoreCase));

            return match?.InternalName;
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

        public string GetTitle(object dataItem)
        {
            try
            {
                SPListItem item = dataItem as SPListItem;
                if (item == null)
                    return string.Empty;

                string configured = string.IsNullOrWhiteSpace(TitleField) ? "Title" : TitleField;
                string resolved = ResolveFieldInternalName(item.ParentList, configured);

                if (!string.IsNullOrEmpty(resolved) && item.Fields.ContainsField(resolved) && item[resolved] != null)
                    return item[resolved].ToString();

                if (item.Fields.ContainsField("Title") && item["Title"] != null)
                    return item["Title"].ToString();

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public string GetImageUrl(object dataItem)
        {
            try
            {
                SPListItem item = dataItem as SPListItem;
                if (item == null || string.IsNullOrEmpty(ImageField))
                    return "/_layouts/15/images/placeholder.png";

                string resolvedImageField = ResolveFieldInternalName(item.ParentList, ImageField);
                if (string.IsNullOrEmpty(resolvedImageField))
                    return "/_layouts/15/images/placeholder.png";

                if (item.Fields.ContainsField(resolvedImageField) && item[resolvedImageField] != null)
                {
                    string fieldValue = item[resolvedImageField].ToString();
                    if (fieldValue.StartsWith("http") || fieldValue.StartsWith("/"))
                        return fieldValue;

                    if (fieldValue.Contains("http") || fieldValue.Contains("/"))
                    {
                        Match match = Regex.Match(fieldValue, @"https?://[^\s\"\",<>]+");
                        if (match.Success)
                            return match.Value;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error getting image URL: " + ex.Message);
            }

            return "/_layouts/15/images/placeholder.png";
        }

        public string FormatDate(object dataItem)
        {
            try
            {
                SPListItem item = dataItem as SPListItem;
                if (item == null || string.IsNullOrEmpty(DateField))
                    return string.Empty;

                string resolvedDateField = ResolveFieldInternalName(item.ParentList, DateField);
                if (string.IsNullOrEmpty(resolvedDateField) || !item.Fields.ContainsField(resolvedDateField) || item[resolvedDateField] == null)
                    return string.Empty;

                object fieldValue = item[resolvedDateField];
                DateTime dateValue;

                if (fieldValue is DateTime dt)
                    dateValue = dt;
                else if (!DateTime.TryParse(fieldValue.ToString(), out dateValue))
                    return string.Empty;

                CultureInfo culture = new CultureInfo("ar-EG");
                string formatted = dateValue.ToString("dd MMMM yyyy", culture);
                return ToArabicIndicDigits(formatted);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static Dictionary<string, string> _descriptionCache = new Dictionary<string, string>();

        public string GetDescription(object dataItem)
        {
            try
            {
                SPListItem item = dataItem as SPListItem;
                if (item == null || string.IsNullOrEmpty(DescriptionField))
                    return string.Empty;

                string cacheKey = $"{item.ID}_{DescriptionField}";
                if (_descriptionCache.TryGetValue(cacheKey, out string cachedDesc))
                    return cachedDesc;

                string resolvedDescriptionField = ResolveFieldInternalName(item.ParentList, DescriptionField);
                if (string.IsNullOrEmpty(resolvedDescriptionField) || !item.Fields.ContainsField(resolvedDescriptionField) || item[resolvedDescriptionField] == null)
                    return string.Empty;

                string descriptionText = item[resolvedDescriptionField].ToString();
                string description = TruncateDescription(descriptionText);
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

            const int maxLength = 150;
            if (descText.Length <= maxLength)
                return descText;

            int lastSpace = descText.Substring(0, maxLength).LastIndexOf(' ');
            if (lastSpace <= 0)
                lastSpace = maxLength;

            return descText.Substring(0, lastSpace) + "...";
        }

        private static Dictionary<string, HyperlinkInfo> _hyperlinkCache = new Dictionary<string, HyperlinkInfo>();

        public HyperlinkInfo GetHyperlinkInfo(object dataItem)
        {
            try
            {
                SPListItem item = dataItem as SPListItem;
                if (item == null || string.IsNullOrEmpty(HyperlinkField))
                    return new HyperlinkInfo { Url = "#", DisplayText = "اقرأ المزيد..." };

                string cacheKey = $"{item.ID}_{HyperlinkField}";
                if (_hyperlinkCache.TryGetValue(cacheKey, out HyperlinkInfo cachedLink))
                    return cachedLink;

                HyperlinkInfo result = new HyperlinkInfo { Url = "#", DisplayText = "اقرأ المزيد..." };

                string resolvedHyperlinkField = ResolveFieldInternalName(item.ParentList, HyperlinkField);
                if (!string.IsNullOrEmpty(resolvedHyperlinkField) && item.Fields.ContainsField(resolvedHyperlinkField) && item[resolvedHyperlinkField] != null)
                {
                    string fieldValue = item[resolvedHyperlinkField].ToString();
                    if (fieldValue.Contains(","))
                    {
                        string[] parts = fieldValue.Split(new char[] { ',' }, 2);
                        if (parts.Length == 2)
                            result = new HyperlinkInfo { Url = parts[0], DisplayText = parts[1] };
                    }
                    else if (fieldValue.StartsWith("http") || fieldValue.StartsWith("/"))
                    {
                        result = new HyperlinkInfo { Url = fieldValue, DisplayText = "اقرأ المزيد..." };
                    }
                }

                _hyperlinkCache[cacheKey] = result;
                return result;
            }
            catch
            {
                return new HyperlinkInfo { Url = "#", DisplayText = "اقرأ المزيد..." };
            }
        }

        private static string ToArabicIndicDigits(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            char[] arabicDigits = new[] { '٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩' };
            StringBuilder sb = new StringBuilder(input.Length);
            foreach (char c in input)
            {
                if (c >= '0' && c <= '9')
                    sb.Append(arabicDigits[c - '0']);
                else
                    sb.Append(c);
            }
            return sb.ToString();
        }
    }

    public class HyperlinkInfo
    {
        public string Url { get; set; }
        public string DisplayText { get; set; }
    }
}
