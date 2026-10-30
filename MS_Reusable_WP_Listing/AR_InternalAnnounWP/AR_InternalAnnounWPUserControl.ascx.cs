using System;
using System.Globalization;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.SharePoint;

namespace MS_Reusable_WP_Listing.AR_InternalAnnounWP
{
    public partial class AR_InternalAnnounWPUserControl : UserControl
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

        private string _detailsPageUrl = "DetailsPage.aspx";
        public string DetailsPageUrl
        {
            get { return _detailsPageUrl; }
            set { _detailsPageUrl = value; }
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
                DataPager pager = lsvImages.FindControl("_dpPager") as DataPager;
                if (pager != null)
                {
                    pager.PageSize = PageSize;
                }

                LoadItems();
            }
        }

        protected void lsvImages_PagePropertiesChanging(object sender, PagePropertiesChangingEventArgs e)
        {
            DataPager pager = lsvImages.FindControl("_dpPager") as DataPager;
            if (pager != null)
            {
                pager.SetPageProperties(e.StartRowIndex, e.MaximumRows, false);
            }

            LoadItems();
        }

        private void LoadItems()
        {
            try
            {
                if (string.IsNullOrEmpty(ListName))
                {
                    ShowError("يرجى تحديد اسم القائمة من خصائص الويب بارت.");
                    return;
                }

                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                {
                    SPWeb web = null;

                    try
                    {
                        if (!string.IsNullOrEmpty(WebUrl))
                        {
                            web = site.OpenWeb(WebUrl);
                        }
                        else
                        {
                            web = site.OpenWeb();
                        }
                    }
                    catch (Exception webEx)
                    {
                        ShowError("خطأ أثناء فتح الموقع. يرجى التحقق من رابط الموقع. " + webEx.Message);
                        return;
                    }

                    if (web == null)
                    {
                        ShowError("تعذر فتح الموقع المحدد. يرجى التحقق من رابط الموقع.");
                        return;
                    }

                    using (web)
                    {
                        SPList list = web.Lists.TryGetList(ListName);
                        if (list == null)
                        {
                            ShowError("القائمة '" + ListName + "' غير موجودة. يرجى التحقق من اسم القائمة.");
                            return;
                        }

                        string resolvedTitleField = ResolveFieldInternalName(list, TitleField);
                        string resolvedDescriptionField = ResolveFieldInternalName(list, DescriptionField);
                        string resolvedImageField = ResolveFieldInternalName(list, ImageField);
                        string resolvedDateField = ResolveFieldInternalName(list, DateField);

                        SPQuery query = new SPQuery();

                        string viewFields = "<FieldRef Name='ID' />";

                        if (!string.IsNullOrEmpty(resolvedTitleField))
                            viewFields += "<FieldRef Name='" + resolvedTitleField + "' />";
                        else
                            viewFields += "<FieldRef Name='Title' />";

                        if (!string.IsNullOrEmpty(resolvedDescriptionField))
                            viewFields += "<FieldRef Name='" + resolvedDescriptionField + "' />";

                        if (!string.IsNullOrEmpty(resolvedImageField))
                            viewFields += "<FieldRef Name='" + resolvedImageField + "' />";

                        if (list.Fields.ContainsField("PublishingPageImage"))
                            viewFields += "<FieldRef Name='PublishingPageImage' />";

                        if (list.Fields.ContainsField("Picture"))
                            viewFields += "<FieldRef Name='Picture' />";

                        if (!string.IsNullOrEmpty(resolvedDateField))
                            viewFields += "<FieldRef Name='" + resolvedDateField + "' />";

                        string[] docLibFields = new[] { "FileRef", "FileDirRef", "FileLeafRef", "EncodedAbsUrl" };
                        foreach (string fieldName in docLibFields)
                        {
                            if (list.Fields.ContainsField(fieldName))
                                viewFields += "<FieldRef Name='" + fieldName + "' />";
                        }

                        query.ViewFields = viewFields;
                        query.ViewFieldsOnly = false;

                        if (!string.IsNullOrEmpty(resolvedDateField))
                            query.Query = "<OrderBy><FieldRef Name='" + resolvedDateField + "' Ascending='False' /></OrderBy>";
                        else
                            query.Query = "<OrderBy><FieldRef Name='ID' Ascending='False' /></OrderBy>";

                        query.Query = ApplyApprovedItemsFilter(list, query.Query);

                        SPListItemCollection items = list.GetItems(query);

                        if (items == null || items.Count == 0)
                        {
                            lsvImages.EmptyDataTemplate = new EmptyDataTemplateHelper("لا توجد عناصر في القائمة.");
                            lsvImages.DataSource = null;
                            lsvImages.DataBind();
                            return;
                        }

                        lsvImages.DataSource = items;
                        lsvImages.DataBind();

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

        private void ShowError(string message)
        {
            litError.Text = message;
            pnlError.Visible = true;
            pnlContent.Visible = false;
            UpdatePanel1?.Update();
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

        public string FormatDate(object dataItem)
        {
            SPListItem item = dataItem as SPListItem;
            if (item == null)
                return string.Empty;

            try
            {
                string configured = DateField;
                string resolved = ResolveFieldInternalName(item.ParentList, configured);

                if (!string.IsNullOrEmpty(resolved) && item.Fields.ContainsField(resolved) && item[resolved] != null)
                {
                    if (DateTime.TryParse(item[resolved].ToString(), out DateTime dateValue))
                    {
                        CultureInfo culture = new CultureInfo("ar-EG");
                        return dateValue.ToString("dd MMMM yyyy", culture);
                    }
                }
            }
            catch
            {
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

                string resolved = ResolveFieldInternalName(item.ParentList, DescriptionField);
                if (!string.IsNullOrEmpty(resolved) && item.Fields.ContainsField(resolved) && item[resolved] != null)
                    return TruncateDescription(item[resolved]);

                string[] commonDescriptionFields = new[] { "Description", "Body", "Content", "Text", "Comments", "Details" };
                foreach (string fieldName in commonDescriptionFields)
                {
                    if (item.Fields.ContainsField(fieldName) && item[fieldName] != null)
                        return TruncateDescription(item[fieldName]);
                }

                if (item.Fields.ContainsField("Title") && item["Title"] != null)
                    return TruncateDescription(item["Title"]);

                return string.Empty;
            }
            catch
            {
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

                int maxLength = 150;
                if (descText.Length <= maxLength)
                    return descText;

                int lastSpace = descText.Substring(0, maxLength).LastIndexOf(' ');
                if (lastSpace <= 0)
                    lastSpace = maxLength;

                return descText.Substring(0, lastSpace) + "...";
            }
            catch
            {
                return string.Empty;
            }
        }
    }

    public class EmptyDataTemplateHelper : ITemplate
    {
        private readonly string _message;

        public EmptyDataTemplateHelper(string message)
        {
            _message = message;
        }

        public void InstantiateIn(Control container)
        {
            Literal literal = new Literal();
            literal.Text = "<div class='no-data' style='direction: rtl; text-align: right;'>" + _message + "</div>";
            container.Controls.Add(literal);
        }
    }
}
