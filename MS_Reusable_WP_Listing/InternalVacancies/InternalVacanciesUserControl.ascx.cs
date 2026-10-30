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

namespace MS_Reusable_WP_Listing.InternalVacancies
{
    public partial class InternalVacanciesUserControl : UserControl
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
                    ShowMessage("Please configure the list name in the web part properties.", true);
                    return;
                }
                
                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                using (SPWeb web = site.OpenWeb())
                {
                    // Check if the list exists
                    SPList list = web.Lists.TryGetList(ListName);
                    if (list == null)
                    {
                        ShowMessage($"The list '{ListName}' does not exist.", true);
                        return;
                    }

                    List<string> displayNamesList = new List<string>();
                    if (!string.IsNullOrEmpty(Columns))
                    {
                        displayNamesList = Columns.Split(',').Select(c => c.Trim()).Where(c => !string.IsNullOrEmpty(c)).ToList();
                    }
                    else
                    {
                        displayNamesList.Add("Title");
                    }

                    Dictionary<string, string> displayToInternalMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    _internalToDisplayMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    displayToInternalMap["ID"] = "ID";
                    _internalToDisplayMap["ID"] = "ID";

                    foreach (SPField field in list.Fields)
                    {
                        if (!field.Hidden)
                        {
                            displayToInternalMap[field.Title] = field.InternalName;
                            _internalToDisplayMap[field.InternalName] = field.Title;
                        }
                    }

                    _pdfInternalName = string.Empty;
                    if (!string.IsNullOrEmpty(PdfField))
                    {
                        if (displayToInternalMap.ContainsKey(PdfField))
                            _pdfInternalName = displayToInternalMap[PdfField];
                        else
                            _pdfInternalName = PdfField;
                    }

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

                    System.Text.StringBuilder viewFieldsBuilder = new System.Text.StringBuilder();
                    viewFieldsBuilder.Append("<FieldRef Name='ID' />");
                    foreach (string col in columnsList)
                        viewFieldsBuilder.Append($"<FieldRef Name='{col}' />");
                    if (!string.IsNullOrEmpty(_pdfInternalName))
                        viewFieldsBuilder.Append($"<FieldRef Name='{_pdfInternalName}' />");

                    SPQuery query = new SPQuery();
                    query.ViewFields = viewFieldsBuilder.ToString();
                    query.ViewFieldsOnly = true;

                    query.Query = ApplyApprovedItemsFilter(list, query.Query);

                    SPListItemCollection items = list.GetItems(query);
                    if (items.Count == 0)
                    {
                        ShowMessage("No items found in the list.", false);
                        return;
                    }

                    DataTable dt = new DataTable();
                    dt.Columns.Add("Id", typeof(int));

                    foreach (string columnInternalName in columnsList)
                    {
                        if (!_internalToDisplayMap.ContainsKey(columnInternalName))
                        {
                            try
                            {
                                SPField field = list.Fields.GetFieldByInternalName(columnInternalName);
                                _internalToDisplayMap[columnInternalName] = field != null ? field.Title : columnInternalName;
                            }
                            catch
                            {
                                _internalToDisplayMap[columnInternalName] = columnInternalName;
                            }
                        }

                        if (items.Count > 0 && items[0][columnInternalName] != null)
                        {
                            object fieldValue = items[0][columnInternalName];
                            Type fieldType = typeof(string);
                            if (fieldValue is DateTime)
                                fieldType = typeof(DateTime);
                            else if (fieldValue is int || fieldValue is long)
                                fieldType = typeof(int);
                            else if (fieldValue is double || fieldValue is decimal)
                                fieldType = typeof(double);
                            else if (fieldValue is bool)
                                fieldType = typeof(bool);

                            if (!dt.Columns.Contains(columnInternalName))
                                dt.Columns.Add(columnInternalName, fieldType);
                        }
                        else
                        {
                            if (!dt.Columns.Contains(columnInternalName))
                                dt.Columns.Add(columnInternalName, typeof(string));
                        }
                    }

                    if (!string.IsNullOrEmpty(_pdfInternalName))
                    {
                        dt.Columns.Add("PdfUrl", typeof(string));
                    }

                    foreach (SPListItem item in items)
                    {
                        DataRow dr = dt.NewRow();
                        dr["Id"] = item.ID;

                        foreach (string column in columnsList)
                        {
                            if (item[column] != null)
                            {
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

                        if (!string.IsNullOrEmpty(_pdfInternalName) && dt.Columns.Contains("PdfUrl"))
                        {
                            string pdfUrl = "#";
                            try
                            {
                                if (item.Fields.ContainsField(_pdfInternalName) && item[_pdfInternalName] != null)
                                {
                                    if (item[_pdfInternalName] is SPFieldUrlValue)
                                    {
                                        SPFieldUrlValue urlValue = new SPFieldUrlValue(item[_pdfInternalName].ToString());
                                        pdfUrl = urlValue.Url;
                                    }
                                    else
                                    {
                                        pdfUrl = item[_pdfInternalName].ToString();
                                    }
                                }
                            }
                            catch
                            {
                                pdfUrl = "#";
                            }

                            dr["PdfUrl"] = string.IsNullOrEmpty(pdfUrl) ? "#" : pdfUrl;
                        }
                        dt.Rows.Add(dr);
                    }

                    _originalData = dt.Copy();
                    _filteredData = _originalData;
                    BindGridView();
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error loading vacancies: {ex.Message}", true);
            }
        }

#if false

        protected void gvVacancies_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gvVacancies.PageIndex = e.NewPageIndex;
            // Don't reload from SharePoint, just rebind the filtered data
            BindGridView();
        }
        
        protected void btnSearch_Click(object sender, EventArgs e)
        {
                            if (cleanName != field.InternalName && !internalToDisplayMap.ContainsKey(cleanName))
                                internalToDisplayMap[cleanName] = field.Title;
                        }
                    }

                    // Add some common mappings for standard SharePoint fields
                    if (!internalToDisplayMap.ContainsKey("Title"))
                        internalToDisplayMap["Title"] = "Title";
                    if (!internalToDisplayMap.ContainsKey("JobTitle"))
                        internalToDisplayMap["JobTitle"] = "Job Title";
                    if (!internalToDisplayMap.ContainsKey("Department"))
                        internalToDisplayMap["Department"] = "Department";
                    if (!internalToDisplayMap.ContainsKey("Location"))
                        internalToDisplayMap["Location"] = "Location";
                    if (!internalToDisplayMap.ContainsKey("ClosingDate"))
                        internalToDisplayMap["ClosingDate"] = "Closing Date";

                    // Convert display names to internal names and validate they exist
                    List<string> columnsList = new List<string>();
                    List<string> validatedColumnsList = new List<string>();

                    foreach (string displayName in displayNamesList)
                    {
                        string internalName = null;

                        if (displayToInternalMap.ContainsKey(displayName))
                        {
                            // Found a match by display name
                            internalName = displayToInternalMap[displayName];
                        }
                        else
                        {
                            // No match found, check if it's a valid internal name
                            foreach (SPField field in list.Fields)
                            {
                                if (field.InternalName.Equals(displayName, StringComparison.OrdinalIgnoreCase))
                                {
                                    internalName = field.InternalName;
                                    break;
                                }
                            }
                        }

                        // Only add if the field actually exists in the list
                        if (internalName != null && list.Fields.ContainsField(internalName))
                        {
                            columnsList.Add(internalName);
                            validatedColumnsList.Add(internalName);
                        }
                    }

                    // If no valid columns were found, add some basic fields that exist
                    if (validatedColumnsList.Count == 0)
                    {
                        foreach (SPField field in list.Fields)
                        {
                            if (!field.Hidden && !field.ReadOnlyField &&
                                field.Type != SPFieldType.Computed &&
                                field.Type != SPFieldType.Recurrence)
                            {
                                columnsList.Add(field.InternalName);
                                validatedColumnsList.Add(field.InternalName);
                                if (validatedColumnsList.Count >= 3) break; // Limit to 3 fields
                            }
                        }
                    }

                    // Always include ID
                    if (!columnsList.Contains("ID") && !columnsList.Contains("Id"))
                    {
                        columnsList.Add("ID");
                    }

                    // Add PDF field if specified
                    string pdfInternalName = string.Empty;
                    if (!string.IsNullOrEmpty(PdfField))
                    {
                        if (displayToInternalMap.ContainsKey(PdfField))
                        {
                            pdfInternalName = displayToInternalMap[PdfField];
                            if (!columnsList.Contains(pdfInternalName))
                            {
                                columnsList.Add(pdfInternalName);
                            }
                        }
                        else
                        {
                            // Try using it as is
                            pdfInternalName = PdfField;
                            if (!columnsList.Contains(pdfInternalName))
                            {
                                System.Text.StringBuilder viewFieldsBuilder = new System.Text.StringBuilder();

                                // Add all fields to the view
                                foreach (SPField field in list.Fields)
                                {
                                    if (!field.Hidden && !field.ReadOnlyField)
                                    {
                                        viewFieldsBuilder.Append($"<FieldRef Name='{field.InternalName}' />");
                                    }
                                }

                                SPQuery query = new SPQuery();
                                query.ViewFields = viewFieldsBuilder.ToString();
                                query.ViewFieldsOnly = true;

                                query.Query = ApplyApprovedItemsFilter(list, query.Query);

                                SPListItemCollection items = list.GetItems(query);

                                if (items.Count == 0)
                                {
                                    ShowMessage("No vacancies found in the list.", false);
                                    return;
                                }

                                // Create DataTable with dynamic columns
                                DataTable dt = new DataTable();
                                dt.Columns.Add("ID", typeof(int));

                                // Add columns from the list
                                foreach (string columnInternalName in columnsList)
                                {
                                    if (columnInternalName.Equals("ID", StringComparison.OrdinalIgnoreCase))
                                        continue; // Already added

                                    if (!dt.Columns.Contains(columnInternalName))
                                    {
                                        // Try to determine the column type from the first item
                                        if (items.Count > 0 && items[0][columnInternalName] != null)
                                        {
                                            object fieldValue = items[0][columnInternalName];
                                            Type fieldType = typeof(string); // Default type

                                            if (fieldValue is DateTime)
                                                fieldType = typeof(DateTime);
                                            else if (fieldValue is int || fieldValue is long)
                                                fieldType = typeof(int);
                                            else if (fieldValue is double || fieldValue is decimal)
                                                fieldType = typeof(double);
                                            else if (fieldValue is bool)
                                                fieldType = typeof(bool);

                                            dt.Columns.Add(columnInternalName, fieldType);
                                        }
                                        else
                                        {
                                            dt.Columns.Add(columnInternalName, typeof(string));
                                        }
                                    }
                                }

                                // Add PdfUrl column for the PDF link
                                dt.Columns.Add("PdfUrl", typeof(string));

                                // Fill the DataTable with list items
                                foreach (SPListItem item in items)
                                {
                                    DataRow dr = dt.NewRow();
                                    dr["ID"] = item.ID;

                                    foreach (string column in columnsList)
                                    {
                                        if (column.Equals("ID", StringComparison.OrdinalIgnoreCase))
                                            continue; // Already added

                                        try
                                        {
                                            if (item.Fields.ContainsField(column) && item[column] != null)
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
                                        catch
                                        {
                                            // If the field doesn't exist, just set it to empty
                                            dr[column] = DBNull.Value;
                                        }
                                    }

                                    // Set the PDF URL
                                    try
                                    {
                                        if (!string.IsNullOrEmpty(pdfInternalName) && item.Fields.ContainsField(pdfInternalName) && item[pdfInternalName] != null)
                                        {
                                            // Handle different field types for PDF URL
                                            if (item[pdfInternalName] is SPFieldUrlValue)
                                            {
                                                SPFieldUrlValue urlValue = new SPFieldUrlValue(item[pdfInternalName].ToString());
                                                dr["PdfUrl"] = urlValue.Url;
                                            }
                                            else
                                            {
                                                dr["PdfUrl"] = item[pdfInternalName].ToString();
                                            }
                                        }
                                        else
                                        {
                                            dr["PdfUrl"] = "#"; // Default empty link
                                        }
                                    }
                                    catch
                                    {
                                        // If the PDF field doesn't exist, just set it to empty
                                        dr["PdfUrl"] = "#";
                                    }

                                    dt.Rows.Add(dr);
                                }

                                // Store the original data
                                _originalData = dt.Copy();

                                // Debug: Log all available columns in the data
                                string availableColumns = "Available columns in data: ";
                                foreach (DataColumn col in dt.Columns)
                                {
                                    availableColumns += col.ColumnName + ", ";
                                }
                                System.Diagnostics.Debug.WriteLine(availableColumns);

                                // Debug: Log the columns specified in the web part properties
                                System.Diagnostics.Debug.WriteLine("Columns specified in web part: " + Columns);

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
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error loading vacancies: {ex.Message}", true);
            }
        }

#endif

        protected void gvVacancies_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gvVacancies.PageIndex = e.NewPageIndex;
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
                    ShowMessage($"Please enter a valid page number between 1 and {totalPages}", true);
                }
            }
            else
            {
                ShowMessage("Please enter a valid page number", true);
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
            
            // Debug: Log the columns in the data
            System.Diagnostics.Debug.WriteLine("Columns in the data:");
            foreach (DataColumn col in _filteredData.Columns)
            {
                System.Diagnostics.Debug.WriteLine($"- {col.ColumnName} ({col.DataType})");
            }
            
            // Create a DataView from the filtered data
            DataView dv = new DataView(_filteredData);
            
            // Set AutoGenerateColumns to false so we can add our own columns with display names
            gvVacancies.AutoGenerateColumns = false;
            
            // Clear existing columns
            gvVacancies.Columns.Clear();
            
            // Add ID column (hidden)
            BoundField idField = new BoundField();
            idField.DataField = "Id";
            idField.HeaderText = "ID";
            idField.Visible = false;
            gvVacancies.Columns.Add(idField);
            
            // Add columns with proper display names
            foreach (DataColumn column in _filteredData.Columns)
            {
                // Skip ID and PdfUrl columns
                if (column.ColumnName == "Id" || column.ColumnName == "PdfUrl")
                    continue;
                
                // Create a new bound field
                BoundField field = new BoundField();
                field.DataField = column.ColumnName;
                
                // Format the display name
                string displayName;
                if (_internalToDisplayMap != null && _internalToDisplayMap.ContainsKey(column.ColumnName))
                    displayName = _internalToDisplayMap[column.ColumnName];
                else
                    displayName = column.ColumnName.Replace("_x0020_", " ").Replace("_", " ");

                field.HeaderText = displayName;
                
                // Format date columns
                if (column.DataType == typeof(DateTime))
                    field.DataFormatString = "{0:MM/dd/yyyy}";
                
                gvVacancies.Columns.Add(field);
            }
            
            // Add PDF column at the end
            if (!string.IsNullOrEmpty(_pdfInternalName) && _filteredData.Columns.Contains("PdfUrl"))
            {
                TemplateField pdfField = new TemplateField();
                pdfField.HeaderText = "PDF";
                pdfField.ItemTemplate = new PdfLinkTemplate();
                gvVacancies.Columns.Add(pdfField);
            }
            
            // Bind the data
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
            
            // Debug: Log the number of columns added
            System.Diagnostics.Debug.WriteLine($"Added {gvVacancies.Columns.Count} columns to GridView");
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
        
        // Custom template class for PDF links
        private class PdfLinkTemplate : ITemplate
        {
            public void InstantiateIn(Control container)
            {
                HyperLink link = new HyperLink();
                link.CssClass = "pdf-link";
                link.Target = "_blank";
                link.Text = "View ";
                
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
                                e.Row.Cells[cellIndex].Text = dateValue.ToString("MM/dd/yyyy");
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
    }
}
