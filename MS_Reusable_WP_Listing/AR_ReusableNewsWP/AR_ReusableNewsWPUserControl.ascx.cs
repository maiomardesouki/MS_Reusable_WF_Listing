using System;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.SharePoint;
using MS_Reusable_WP_Listing.ReusableNewsWP;

namespace MS_Reusable_WP_Listing.AR_ReusableNewsWP
{
    public partial class AR_ReusableNewsWPUserControl : UserControl
    {
        public string ListName { get; set; }
        public string WebUrl { get; set; }
        public string TitleField { get; set; }
        public string DescriptionField { get; set; }
        public string DateField { get; set; }
        public string LocationField { get; set; }
        public string ImageField { get; set; }
        public string DetailsPageUrl { get; set; }
        public int PageSize { get; set; }

        private Stopwatch stopwatch;
        private const int MaxExecutionTimeMs = 10000;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
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
                if (string.IsNullOrEmpty(ListName))
                {
                    ShowError("يرجى تحديد اسم القائمة من خصائص الويب بارت.");
                    return;
                }

                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                using (SPWeb web = string.IsNullOrEmpty(WebUrl) ? site.OpenWeb() : site.OpenWeb(WebUrl))
                {
                    SPList list = web.Lists.TryGetList(ListName);
                    if (list == null)
                    {
                        ShowError("القائمة '" + ListName + "' غير موجودة.");

                        StringBuilder listsLog = new StringBuilder();
                        listsLog.AppendLine("Available lists:");
                        foreach (SPList availableList in web.Lists)
                        {
                            listsLog.AppendLine(" - " + availableList.Title);
                        }
                        Debug.WriteLine(listsLog.ToString());

                        return;
                    }

                    string resolvedTitleField = ResolveFieldInternalName(list, TitleField);
                    string resolvedDescriptionField = ResolveFieldInternalName(list, DescriptionField);
                    string resolvedDateField = ResolveFieldInternalName(list, DateField);
                    string resolvedLocationField = ResolveFieldInternalName(list, LocationField);
                    string resolvedImageField = ResolveFieldInternalName(list, ImageField);

                    StringBuilder viewFieldsBuilder = new StringBuilder();
                    viewFieldsBuilder.Append("<FieldRef Name='ID' />");
                    viewFieldsBuilder.Append("<FieldRef Name='Title' />");

                    if (!string.IsNullOrEmpty(resolvedTitleField))
                        viewFieldsBuilder.Append("<FieldRef Name='" + resolvedTitleField + "' />");

                    if (!string.IsNullOrEmpty(resolvedDescriptionField))
                        viewFieldsBuilder.Append("<FieldRef Name='" + resolvedDescriptionField + "' />");

                    if (!string.IsNullOrEmpty(resolvedDateField))
                        viewFieldsBuilder.Append("<FieldRef Name='" + resolvedDateField + "' />");

                    if (!string.IsNullOrEmpty(resolvedLocationField))
                        viewFieldsBuilder.Append("<FieldRef Name='" + resolvedLocationField + "' />");

                    if (!string.IsNullOrEmpty(resolvedImageField))
                        viewFieldsBuilder.Append("<FieldRef Name='" + resolvedImageField + "' />");

                    SPQuery query = new SPQuery();
                    query.ViewFields = viewFieldsBuilder.ToString();
                    query.ViewFieldsOnly = true;
                    query.RowLimit = (uint)PageSize;

                    if (!string.IsNullOrEmpty(resolvedDateField))
                        query.Query = "<OrderBy><FieldRef Name='" + resolvedDateField + "' Ascending='False' /></OrderBy>";

                    query.Query = ApplyApprovedItemsFilter(list, query.Query);

                    SPListItemCollection items = list.GetItems(query);

                    if (items.Count == 0)
                    {
                        ShowError("لا توجد أخبار.", false);
                        return;
                    }

                    DataTable dt = new DataTable();
                    dt.Columns.Add("ID", typeof(int));
                    dt.Columns.Add("Title", typeof(string));
                    dt.Columns.Add("Description", typeof(string));
                    dt.Columns.Add("DateDay", typeof(string));
                    dt.Columns.Add("DateMonth", typeof(string));
                    dt.Columns.Add("Location", typeof(string));
                    dt.Columns.Add("ImageUrl", typeof(string));
                    dt.Columns.Add("DetailsUrl", typeof(string));

                    foreach (SPListItem item in items)
                    {
                        try
                        {
                            DataRow row = dt.NewRow();
                            row["ID"] = item.ID;
                            row["Title"] = GetTitle(item, resolvedTitleField);
                            row["Description"] = GetDescription(item, resolvedDescriptionField);
                            row["DateDay"] = FormatDateDay(item, resolvedDateField);
                            row["DateMonth"] = FormatDateMonth(item, resolvedDateField);
                            row["Location"] = GetLocation(item, resolvedLocationField);
                            row["ImageUrl"] = GetImageUrl(item, resolvedImageField);
                            row["DetailsUrl"] = GetDetailsUrl(item);
                            dt.Rows.Add(row);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine("Error processing item " + item.ID + ": " + ex.Message);
                        }
                    }

                    if (dt.Rows.Count > 0)
                    {
                        lsvNews.DataSource = dt;
                        lsvNews.DataBind();
                        pnlContent.Visible = true;
                        pnlError.Visible = false;
                    }
                    else
                    {
                        ShowError("لا توجد أخبار بالحقول المحددة.", false);
                    }

                    if (stopwatch.ElapsedMilliseconds > MaxExecutionTimeMs)
                    {
                        ShowError("انتهت مهلة العملية. يرجى المحاولة مرة أخرى أو التواصل مع المسؤول.");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError("حدث خطأ أثناء تحميل الأخبار: " + ex.Message);
                Debug.WriteLine("Error in AR_ReusableNewsWP: " + ex);
            }
            finally
            {
                stopwatch.Stop();
            }
        }

        protected void lsvNews_PagePropertiesChanging(object sender, PagePropertiesChangingEventArgs e)
        {
            DataPager pager = lsvNews.FindControl("_dpPager") as DataPager;
            if (pager != null)
            {
                pager.SetPageProperties(e.StartRowIndex, e.MaximumRows, false);
                LoadItems();
            }
        }

        private void ShowError(string message, bool isError = true)
        {
            litError.Text = message;
            pnlError.Visible = true;
            pnlContent.Visible = !isError;
            pnlError.CssClass = isError ? "error-message" : "info-message";
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

        private string GetTitle(SPListItem item, string resolvedTitleField)
        {
            if (item == null)
                return string.Empty;

            try
            {
                if (!string.IsNullOrEmpty(resolvedTitleField) && item.Fields.ContainsField(resolvedTitleField) && item[resolvedTitleField] != null)
                    return item[resolvedTitleField].ToString();

                if (item.Fields.ContainsField("Title") && item["Title"] != null)
                    return item["Title"].ToString();

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetDescription(SPListItem item, string resolvedDescriptionField)
        {
            if (item == null)
                return string.Empty;

            try
            {
                if (!string.IsNullOrEmpty(resolvedDescriptionField) && item.Fields.ContainsField(resolvedDescriptionField) && item[resolvedDescriptionField] != null)
                {
                    string desc = item[resolvedDescriptionField].ToString();
                    if (desc.Contains("<") && desc.Contains(">"))
                        desc = Regex.Replace(desc, @"<[^>]+>", string.Empty);
                    return desc;
                }

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string FormatDateDay(SPListItem item, string resolvedDateField)
        {
            if (item == null || string.IsNullOrEmpty(resolvedDateField))
                return string.Empty;

            try
            {
                object value = item[resolvedDateField];
                if (value == null)
                    return string.Empty;

                DateTime date;
                if (value is DateTime)
                    date = (DateTime)value;
                else if (!DateTime.TryParse(value.ToString(), out date))
                    return string.Empty;

                return ToArabicIndicDigits(date.Day.ToString());
            }
            catch
            {
                return string.Empty;
            }
        }

        private string FormatDateMonth(SPListItem item, string resolvedDateField)
        {
            if (item == null || string.IsNullOrEmpty(resolvedDateField))
                return string.Empty;

            try
            {
                object value = item[resolvedDateField];
                if (value == null)
                    return string.Empty;

                DateTime date;
                if (value is DateTime)
                    date = (DateTime)value;
                else if (!DateTime.TryParse(value.ToString(), out date))
                    return string.Empty;

                CultureInfo culture = new CultureInfo("ar-EG");
                return date.ToString("MMMM", culture);
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetLocation(SPListItem item, string resolvedLocationField)
        {
            if (item == null || string.IsNullOrEmpty(resolvedLocationField))
                return string.Empty;

            try
            {
                if (item.Fields.ContainsField(resolvedLocationField) && item[resolvedLocationField] != null)
                    return item[resolvedLocationField].ToString();
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetImageUrl(SPListItem item, string resolvedImageField)
        {
            string fieldToUse = resolvedImageField;
            if (string.IsNullOrEmpty(fieldToUse))
                fieldToUse = ImageField;

            return ImageHelper.GetImageUrl(item, fieldToUse);
        }

        private string GetDetailsUrl(SPListItem item)
        {
            if (item == null)
                return "#";

            return DetailsPageUrl + "?aID=" + item.ID;
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
}
