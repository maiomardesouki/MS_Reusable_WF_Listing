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

namespace MS_Reusable_WP_Listing.VisualWebPart1
{
    public partial class VisualWebPart1UserControl : UserControl
    {
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
        
        // Property to store the search term
        private string _searchTerm = string.Empty;
        
        // Property to store filtered data
        private DataTable _filteredData;
        
        // Property to store original data
        private DataTable _originalData;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadEvents();
            }
            else
            {
                if (_originalData == null)
                {
                    _searchTerm = txtSearch == null ? string.Empty : txtSearch.Text.Trim();
                    LoadEvents();
                }
            }
        }

        private void LoadEvents()
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

                    // Parse the columns from the property
                    List<string> displayNamesList = new List<string>();
                    if (!string.IsNullOrEmpty(Columns))
                    {
                        displayNamesList = Columns.Split(',').Select(c => c.Trim()).ToList();
                    }
                    else
                    {
                        // Default columns if none specified
                        displayNamesList.Add("Title");
                    }
                    
                    // Create a mapping between display names and internal names
                    Dictionary<string, string> displayToInternalMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    Dictionary<string, string> internalToDisplayMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    
                    // Always include ID
                    displayToInternalMap["ID"] = "ID";
                    internalToDisplayMap["ID"] = "ID";
                    
                    // Build the mapping for all fields in the list
                    foreach (SPField field in list.Fields)
                    {
                        if (!field.Hidden)
                        {
                            displayToInternalMap[field.Title] = field.InternalName;
                            internalToDisplayMap[field.InternalName] = field.Title;
                        }
                    }
                    
                    // Convert display names to internal names
                    List<string> columnsList = new List<string>();
                    foreach (string displayName in displayNamesList)
                    {
                        if (displayToInternalMap.ContainsKey(displayName))
                        {
                            // Found a match by display name
                            columnsList.Add(displayToInternalMap[displayName]);
                        }
                        else
                        {
                            // No match found, try using it as is (might be an internal name)
                            columnsList.Add(displayName);
                        }
                    }
                    
                    // Always include ID
                    if (!columnsList.Contains("ID") && !columnsList.Contains("Id"))
                    {
                        columnsList.Add("ID");
                    }

                    // Build the ViewFields XML
                    System.Text.StringBuilder viewFieldsBuilder = new System.Text.StringBuilder();
                    foreach (string column in columnsList)
                    {
                        viewFieldsBuilder.Append($"<FieldRef Name='{column}' />");
                    }

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

                    // Create DataTable with dynamic columns
                    DataTable dt = new DataTable();
                    dt.Columns.Add("Id", typeof(int));
                    
                    // We already have the mapping in internalToDisplayMap, but we need to make sure
                    // we have an entry for each column in our list
                    foreach (string columnInternalName in columnsList)
                    {
                        if (columnInternalName.Equals("ID", StringComparison.OrdinalIgnoreCase))
                            continue; // Already added
                            
                        if (!internalToDisplayMap.ContainsKey(columnInternalName))
                        {
                            // Try to get the display name from the field
                            try
                            {
                                SPField field = list.Fields.GetFieldByInternalName(columnInternalName);
                                if (field != null)
                                {
                                    internalToDisplayMap[columnInternalName] = field.Title;
                                }
                                else
                                {
                                    // If field doesn't exist, just use the internal name
                                    internalToDisplayMap[columnInternalName] = columnInternalName;
                                }
                            }
                            catch
                            {
                                // If field doesn't exist, just use the internal name
                                internalToDisplayMap[columnInternalName] = columnInternalName;
                            }
                        }
                        
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

                            if (!dt.Columns.Contains(columnInternalName))
                                dt.Columns.Add(columnInternalName, fieldType);
                        }
                        else
                        {
                            if (!dt.Columns.Contains(columnInternalName))
                                dt.Columns.Add(columnInternalName, typeof(string));
                        }
                    }

                    // Fill the DataTable with list items
                    foreach (SPListItem item in items)
                    {
                        DataRow dr = dt.NewRow();
                        dr["Id"] = item.ID;

                        foreach (string column in columnsList)
                        {
                            if (column.Equals("ID", StringComparison.OrdinalIgnoreCase))
                                continue; // Already added

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
                    
                    // Update column headers to use display names and hide the ID column
                    if (gvEvents.HeaderRow != null)
                    {                        
                        // Update column headers with display names
                        for (int i = 0; i < gvEvents.HeaderRow.Cells.Count; i++)
                        {
                            string headerText = gvEvents.HeaderRow.Cells[i].Text;
                            if (internalToDisplayMap.ContainsKey(headerText))
                            {
                                gvEvents.HeaderRow.Cells[i].Text = internalToDisplayMap[headerText];
                            }
                        }
                        int idColumnIndex = -1;
                        for (int i = 0; i < gvEvents.HeaderRow.Cells.Count; i++)
                        {
                            if (gvEvents.HeaderRow.Cells[i].Text.Equals("Id", StringComparison.OrdinalIgnoreCase))
                            {
                                idColumnIndex = i;
                                break;
                            }
                        }

                        if (idColumnIndex >= 0)
                        {
                            gvEvents.HeaderRow.Cells[idColumnIndex].Visible = false;
                            foreach (GridViewRow row in gvEvents.Rows)
                            {
                                row.Cells[idColumnIndex].Visible = false;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"Error loading events: {ex.Message}", true);
            }
        }

        protected void gvEvents_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gvEvents.PageIndex = e.NewPageIndex;

            if (_filteredData == null)
            {
                _searchTerm = txtSearch == null ? string.Empty : txtSearch.Text.Trim();
                LoadEvents();
                return;
            }

            // Don't reload from SharePoint, just rebind the filtered data
            BindGridView();
        }
        
        protected void btnSearch_Click(object sender, EventArgs e)
        {
            _searchTerm = txtSearch.Text.Trim();
            gvEvents.PageIndex = 0; // Reset to first page when searching
            
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
                LoadEvents();
            }
        }
        
        protected void btnGoToPage_Click(object sender, EventArgs e)
        {
            if (_filteredData == null)
            {
                _searchTerm = txtSearch == null ? string.Empty : txtSearch.Text.Trim();
                LoadEvents();
            }

            if (_filteredData == null)
                return;

            int pageNumber;
            if (int.TryParse(txtPageNumber.Text, out pageNumber))
            {
                // Adjust for 0-based index
                pageNumber = pageNumber - 1;
                
                // Validate page number
                int totalPages = (int)Math.Ceiling((double)_filteredData.Rows.Count / gvEvents.PageSize);
                if (pageNumber >= 0 && pageNumber < totalPages)
                {
                    gvEvents.PageIndex = pageNumber;
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
            
            // Bind to GridView
            gvEvents.DataSource = dv;
            gvEvents.DataBind();
            
            // Update pagination info
            int totalRecords = _filteredData.Rows.Count;
            int totalPages = (int)Math.Ceiling((double)totalRecords / gvEvents.PageSize);
            
            lblCurrentPage.Text = (gvEvents.PageIndex + 1).ToString();
            lblTotalPages.Text = totalPages.ToString();
            lblTotalRecords.Text = totalRecords.ToString();
            
            // Set the maximum value for the page number input
            txtPageNumber.Attributes["max"] = totalPages.ToString();
        }

        protected void gvEvents_RowDataBound(object sender, GridViewRowEventArgs e)
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
                    
                    // Hide ID column
                    if (column.ColumnName.Equals("Id", StringComparison.OrdinalIgnoreCase))
                    {
                        int idColumnIndex = drv.Row.Table.Columns.IndexOf(column.ColumnName);
                        if (idColumnIndex >= 0 && idColumnIndex < e.Row.Cells.Count)
                        {
                            e.Row.Cells[idColumnIndex].Visible = false;
                        }
                    }
                }
            }
            else if (e.Row.RowType == DataControlRowType.Header)
            {
                // Apply header styling
                e.Row.CssClass = "brown-bg text-white";
                
                // Hide ID column in header
                for (int i = 0; i < e.Row.Cells.Count; i++)
                {
                    if (e.Row.Cells[i].Text.Equals("Id", StringComparison.OrdinalIgnoreCase))
                    {
                        e.Row.Cells[i].Visible = false;
                        break;
                    }
                }
            }
        }


        private void ShowMessage(string message, bool isError)
        {
            ltr_Msg.Text = message;
            pnlMsg.Visible = true;
            pnlMsg.CssClass = isError ? "note-message error-color d-flex w-100" : "note-message info-color d-flex w-100";
        }

        public string GetResourceValue(string key)
        {
            try
            {
                return SPUtility.GetLocalizedString("$Resources:CBE-Resources," + key, "CBE-Resources", (uint)System.Threading.Thread.CurrentThread.CurrentUICulture.LCID);
            }
            catch
            {
                return key;
            }
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
