using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.Utilities;
using System.Globalization;

namespace MS_Reusable_WP_Listing.ARInternalVacancies
{
    public partial class ARInternalVacanciesUserControl : UserControl
    {
        #region Properties
        
        // Property to store the list name from the web part properties
        private string _listName = string.Empty;
        public string ListName
        {
            get { return _listName; }
            set { _listName = value; }
        }

        // Property to store the columns to display
        private string _columns = string.Empty;
        public string Columns
        {
            get { return _columns; }
            set { _columns = value; }
        }
        
        // Property to store the column display names (Arabic)
        private string _columnDisplayNames = string.Empty;
        public string ColumnDisplayNames
        {
            get { return _columnDisplayNames; }
            set { _columnDisplayNames = value; }
        }
        
        // Property to store the PDF field name
        private string _pdfField = string.Empty;
        public string PdfField
        {
            get { return _pdfField; }
            set { _pdfField = value; }
        }
        
        // Property to store the search term
        private string _searchTerm = string.Empty;
        
        // Property to store filtered data
        private DataTable _filteredData;
        
        // Property to store original data
        private DataTable _originalData;

        private Dictionary<string, string> _internalToDisplayMap;
        private Dictionary<string, string> _columnDisplayNameMap;
        private string _pdfInternalName;
        
        #endregion

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadVacancies();
            }
            else
            {
                if (_originalData == null)
                {
                    _searchTerm = txtSearch == null ? string.Empty : txtSearch.Text.Trim();
                    LoadVacancies();
                }
            }
        }

        private void LoadVacancies()
        {
            try
            {
                // Check if list name is provided
                if (string.IsNullOrEmpty(ListName))
                {
                    ShowMessage("الرجاء تكوين اسم القائمة في خصائص جزء الويب.", true);
                    return;
                }

                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                using (SPWeb web = site.OpenWeb())
                {
                    // Check if the list exists
                    SPList list = web.Lists.TryGetList(ListName);
                    if (list == null)
                    {
                        ShowMessage($"القائمة '{ListName}' غير موجودة.", true);
                        return;
                    }

                    // Parse the columns from the property
                    List<string> displayNamesList = new List<string>();
                    if (!string.IsNullOrEmpty(Columns))
                    {
                        displayNamesList = Columns.Split(',').Select(c => c.Trim()).Where(c => !string.IsNullOrEmpty(c)).ToList();
                    }
                    else
                    {
                        // Default columns if none specified
                        displayNamesList.Add("Title");
                        displayNamesList.Add("Department");
                        displayNamesList.Add("Location");
                        displayNamesList.Add("ClosingDate");
                    }
                    
                    // Parse the column display names (Arabic)
                    _columnDisplayNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    if (!string.IsNullOrEmpty(ColumnDisplayNames))
                    {
                        string[] displayNamePairs = ColumnDisplayNames.Split(',');
                        foreach (string pair in displayNamePairs)
                        {
                            string[] parts = pair.Split(':');
                            if (parts.Length == 2)
                            {
                                string columnName = parts[0].Trim();
                                string displayName = parts[1].Trim();
                                if (!string.IsNullOrEmpty(columnName) && !string.IsNullOrEmpty(displayName))
                                    _columnDisplayNameMap[columnName] = displayName;
                            }
                        }
                    }
                    else
                    {
                        // Default Arabic display names if none specified
                        _columnDisplayNameMap["Title"] = "عنوان الوظيفة";
                        _columnDisplayNameMap["Department"] = "القسم";
                        _columnDisplayNameMap["Location"] = "الموقع";
                        _columnDisplayNameMap["ClosingDate"] = "تاريخ الإغلاق";
                    }
                    
                    // Create a mapping between display names and internal names
                    Dictionary<string, string> displayToInternalMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    _internalToDisplayMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    
                    // Always include ID
                    displayToInternalMap["ID"] = "ID";
                    _internalToDisplayMap["ID"] = "ID";
                    
                    // Build the mapping for all fields in the list
                    foreach (SPField field in list.Fields)
                    {
                        if (!field.Hidden)
                        {
                            displayToInternalMap[field.Title] = field.InternalName;
                            _internalToDisplayMap[field.InternalName] = field.Title;
                        }
                    }

                    // Add PDF field if specified
                    _pdfInternalName = string.Empty;
                    if (!string.IsNullOrEmpty(PdfField))
                    {
                        if (displayToInternalMap.ContainsKey(PdfField))
                            _pdfInternalName = displayToInternalMap[PdfField];
                        else
                            _pdfInternalName = PdfField;
                    }

                    // Convert display names to internal names
                    List<string> columnsList = new List<string>();
                    foreach (string displayName in displayNamesList)
                    {
                        string internalName;
                        if (displayToInternalMap.ContainsKey(displayName))
                            internalName = displayToInternalMap[displayName];
                        else
                            internalName = displayName;

                        if (internalName.Equals("ID", StringComparison.OrdinalIgnoreCase))
                            continue;
                        if (!string.IsNullOrEmpty(_pdfInternalName) && internalName.Equals(_pdfInternalName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        if (!columnsList.Contains(internalName, StringComparer.OrdinalIgnoreCase))
                            columnsList.Add(internalName);
                    }

                    // Build the ViewFields XML
                    System.Text.StringBuilder viewFieldsBuilder = new System.Text.StringBuilder();
                    viewFieldsBuilder.Append("<FieldRef Name='ID' />");
                    foreach (string column in columnsList)
                    {
                        viewFieldsBuilder.Append($"<FieldRef Name='{column}' />");
                    }
                    if (!string.IsNullOrEmpty(_pdfInternalName))
                    {
                        viewFieldsBuilder.Append($"<FieldRef Name='{_pdfInternalName}' />");
                    }

                    SPQuery query = new SPQuery();
                    query.ViewFields = viewFieldsBuilder.ToString();
                    query.ViewFieldsOnly = true;

                    query.Query = ApplyApprovedItemsFilter(list, query.Query);
                    
                    SPListItemCollection items = list.GetItems(query);

                    if (items.Count == 0)
                    {
                        ShowMessage("لم يتم العثور على وظائف شاغرة في القائمة.", false);
                        return;
                    }

                    // Create DataTable with dynamic columns
                    DataTable dt = new DataTable();
                    dt.Columns.Add("Id", typeof(int));
                    
                    // Add columns from the list
                    foreach (string column in columnsList)
                    {
                        if (!dt.Columns.Contains(column))
                        {
                            // Try to determine the column type from the first item
                            if (items.Count > 0 && items[0][column] != null)
                            {
                                object fieldValue = items[0][column];
                                Type fieldType = typeof(string); // Default type

                                if (fieldValue is DateTime)
                                    fieldType = typeof(DateTime);
                                else if (fieldValue is int || fieldValue is long)
                                    fieldType = typeof(int);
                                else if (fieldValue is double || fieldValue is decimal)
                                    fieldType = typeof(double);
                                else if (fieldValue is bool)
                                    fieldType = typeof(bool);

                                dt.Columns.Add(column, fieldType);
                            }
                            else
                            {
                                dt.Columns.Add(column, typeof(string));
                            }
                        }
                    }

                    // Add PdfUrl column for the PDF link
                    if (!string.IsNullOrEmpty(_pdfInternalName))
                    {
                        dt.Columns.Add("PdfUrl", typeof(string));
                    }

                    // Fill the DataTable with list items
                    foreach (SPListItem item in items)
                    {
                        DataRow dr = dt.NewRow();
                        dr["Id"] = item.ID;

                        foreach (string column in columnsList)
                        {
                            if (item[column] != null)
                            {
                                // Handle different field types
                                if (item[column] is SPFieldLookupValue)
                                {
                                    SPFieldLookupValue lookupValue = new SPFieldLookupValue(item[column].ToString());
                                    dr[column] = lookupValue.LookupValue;
                                }
                                else if (item[column] is SPFieldUserValue)
                                {
                                    SPFieldUserValue userValue = new SPFieldUserValue(web, item[column].ToString());
                                    dr[column] = userValue.LookupValue;
                                }
                                else if (item[column] is DateTime)
                                {
                                    dr[column] = (DateTime)item[column];
                                }
                                else
                                {
                                    dr[column] = item[column].ToString();
                                }
                            }
                            else
                            {
                                dr[column] = DBNull.Value;
                            }
                        }
                        
                        // Set the PDF URL
                        if (!string.IsNullOrEmpty(_pdfInternalName) && dt.Columns.Contains("PdfUrl"))
                        {
                            if (item.Fields.ContainsField(_pdfInternalName) && item[_pdfInternalName] != null)
                            {
                                // Handle different field types for PDF URL
                                if (item[_pdfInternalName] is SPFieldUrlValue)
                                {
                                    SPFieldUrlValue urlValue = new SPFieldUrlValue(item[_pdfInternalName].ToString());
                                    dr["PdfUrl"] = urlValue.Url;
                                }
                                else
                                {
                                    dr["PdfUrl"] = item[_pdfInternalName].ToString();
                                }
                            }
                            else
                            {
                                dr["PdfUrl"] = "#"; // Default empty link
                            }
                        }

                        dt.Rows.Add(dr);
                    }

                    // Store the original data
                    _originalData = dt.Copy();
                    
                    // Apply search filter if search term is provided
                    if (!string.IsNullOrEmpty(_searchTerm))
                    {
                        ApplySearchFilter();
                    }
                    else
                    {
                        _filteredData = _originalData;
                    }

                    // Bind the filtered data to the GridView
                    BindGridView();
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"خطأ في تحميل الوظائف الشاغرة: {ex.Message}", true);
            }
        }

        protected void gvVacancies_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gvVacancies.PageIndex = e.NewPageIndex;

            if (_filteredData == null)
            {
                _searchTerm = txtSearch == null ? string.Empty : txtSearch.Text.Trim();
                LoadVacancies();
                return;
            }

            // Don't reload from SharePoint, just rebind the filtered data
            BindGridView();
        }
        
        protected void btnSearch_Click(object sender, EventArgs e)
        {
            _searchTerm = txtSearch.Text.Trim();
            gvVacancies.PageIndex = 0; // Reset to first page when searching
            
            if (_originalData != null)
            {
                if (!string.IsNullOrEmpty(_searchTerm))
                {
                    ApplySearchFilter();
                }
                else
                {
                    _filteredData = _originalData;
                }
                
                BindGridView();
            }
            else
            {
                LoadVacancies();
            }
        }
        
        protected void btnGoToPage_Click(object sender, EventArgs e)
        {
            if (_filteredData == null)
            {
                _searchTerm = txtSearch == null ? string.Empty : txtSearch.Text.Trim();
                LoadVacancies();
            }

            if (_filteredData == null)
                return;

            int pageNumber;
            if (int.TryParse(txtPageNumber.Text, out pageNumber))
            {
                // Adjust for 0-based index
                pageNumber = pageNumber - 1;
                
                // Validate page number
                int totalPages = (int)Math.Ceiling((double)_filteredData.Rows.Count / gvVacancies.PageSize);
                if (pageNumber >= 0 && pageNumber < totalPages)
                {
                    gvVacancies.PageIndex = pageNumber;
                    BindGridView();
                }
                else
                {
                    // Invalid page number, show error
                    ShowMessage($"الرجاء إدخال رقم صفحة صالح بين 1 و {totalPages}", true);
                }
            }
            else
            {
                ShowMessage("الرجاء إدخال رقم صفحة صالح", true);
            }
        }
        
        private void ApplySearchFilter()
        {
            if (_originalData == null || string.IsNullOrEmpty(_searchTerm))
                return;
            
            _filteredData = _originalData.Clone();
            
            // Search across all columns
            foreach (DataRow row in _originalData.Rows)
            {
                bool matchFound = false;
                
                foreach (DataColumn column in _originalData.Columns)
                {
                    // Skip the PdfUrl column in search
                    if (column.ColumnName == "PdfUrl")
                        continue;
                        
                    if (row[column] != null && row[column] != DBNull.Value)
                    {
                        string cellValue = row[column].ToString().ToLower();
                        if (cellValue.Contains(_searchTerm.ToLower()))
                        {
                            matchFound = true;
                            break;
                        }
                    }
                }
                
                if (matchFound)
                {
                    _filteredData.ImportRow(row);
                }
            }
        }
        
        private void BindGridView()
        {
            if (_filteredData == null)
                return;
            
            // Create a DataView from the filtered data
            DataView dv = new DataView(_filteredData);

            gvVacancies.AutoGenerateColumns = false;
            gvVacancies.Columns.Clear();

            BoundField idField = new BoundField();
            idField.DataField = "Id";
            idField.HeaderText = "ID";
            idField.Visible = false;
            gvVacancies.Columns.Add(idField);

            foreach (DataColumn column in _filteredData.Columns)
            {
                if (column.ColumnName.Equals("Id", StringComparison.OrdinalIgnoreCase) || column.ColumnName.Equals("PdfUrl", StringComparison.OrdinalIgnoreCase))
                    continue;

                BoundField field = new BoundField();
                field.DataField = column.ColumnName;

                string headerText = column.ColumnName;
                if (_columnDisplayNameMap != null)
                {
                    if (_columnDisplayNameMap.ContainsKey(column.ColumnName))
                        headerText = _columnDisplayNameMap[column.ColumnName];
                    else if (_internalToDisplayMap != null && _internalToDisplayMap.ContainsKey(column.ColumnName) && _columnDisplayNameMap.ContainsKey(_internalToDisplayMap[column.ColumnName]))
                        headerText = _columnDisplayNameMap[_internalToDisplayMap[column.ColumnName]];
                }

                if (headerText == column.ColumnName && _internalToDisplayMap != null && _internalToDisplayMap.ContainsKey(column.ColumnName))
                    headerText = _internalToDisplayMap[column.ColumnName];

                field.HeaderText = headerText;

                if (column.DataType == typeof(DateTime))
                    field.DataFormatString = "{0:dd/MM/yyyy}";

                gvVacancies.Columns.Add(field);
            }

            if (!string.IsNullOrEmpty(_pdfInternalName) && _filteredData.Columns.Contains("PdfUrl"))
            {
                TemplateField pdfField = new TemplateField();
                pdfField.HeaderText = "ملف PDF";
                pdfField.ItemTemplate = new PdfLinkTemplate();
                gvVacancies.Columns.Add(pdfField);
            }

            gvVacancies.DataSource = dv;
            gvVacancies.DataBind();
            
            // Update pagination info
            int totalRecords = _filteredData.Rows.Count;
            int totalPages = (int)Math.Ceiling((double)totalRecords / gvVacancies.PageSize);
            
            lblCurrentPage.Text = (gvVacancies.PageIndex + 1).ToString();
            lblTotalPages.Text = totalPages.ToString();
            lblTotalRecords.Text = totalRecords.ToString();
            
            // Set the maximum value for the page number input
            txtPageNumber.Attributes["max"] = totalPages.ToString();
        }

        private class PdfLinkTemplate : ITemplate
        {
            public void InstantiateIn(Control container)
            {
                HyperLink link = new HyperLink();
                link.CssClass = "pdf-link";
                link.Target = "_blank";
                link.Text = "عرض ";

                Literal icon = new Literal();
                icon.Text = "<i class=\"fas fa-file-pdf\"></i>";

                link.Controls.Add(icon);
                container.Controls.Add(link);

                container.DataBinding += new EventHandler(OnDataBinding);
            }

            private void OnDataBinding(object sender, EventArgs e)
            {
                Control container = (Control)sender;
                IDataItemContainer dataItem = (IDataItemContainer)container.NamingContainer;

                HyperLink link = (HyperLink)container.Controls[0];
                object raw = DataBinder.Eval(dataItem.DataItem, "PdfUrl");
                string url = raw == null ? string.Empty : raw.ToString();
                link.NavigateUrl = string.IsNullOrEmpty(url) ? "#" : url;
                link.Visible = !(string.IsNullOrEmpty(url) || url == "#");
            }
        }

        protected void gvVacancies_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                // Apply formatting to date columns
                DataRowView drv = (DataRowView)e.Row.DataItem;
                foreach (DataColumn column in drv.Row.Table.Columns)
                {
                    if (column.DataType == typeof(DateTime))
                    {
                        int cellIndex = drv.Row.Table.Columns.IndexOf(column.ColumnName);
                        if (cellIndex >= 0 && cellIndex < e.Row.Cells.Count)
                        {
                            DateTime dateValue;
                            if (DateTime.TryParse(e.Row.Cells[cellIndex].Text, out dateValue))
                            {
                                // Format date in Arabic style
                                e.Row.Cells[cellIndex].Text = dateValue.ToString("dd/MM/yyyy");
                            }
                        }
                    }
                }
            }
            else if (e.Row.RowType == DataControlRowType.Header)
            {
                // Apply header styling
                e.Row.CssClass = "brown-bg text-white";
            }
        }

        private void ShowMessage(string message, bool isError)
        {
            ltr_Msg.Text = message;
            pnlMsg.Visible = true;
            pnlMsg.CssClass = isError ? "note-message error-color d-flex w-100" : "note-message info-color d-flex w-100";
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
}
