using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.SharePoint;

namespace MS_Reusable_WP_Listing.AR_EmployeeHandbook
{
    public partial class AR_EmployeeHandbookUserControl : UserControl
    {
        private string _listName = string.Empty;
        public string ListName
        {
            get { return _listName; }
            set { _listName = value; }
        }

        private string _columns = string.Empty;
        public string Columns
        {
            get { return _columns; }
            set { _columns = value; }
        }

        private string _searchTerm = string.Empty;
        private DataTable _filteredData;
        private DataTable _originalData;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadItems();
            }
            else
            {
                if (_originalData == null)
                {
                    _searchTerm = txtSearch == null ? string.Empty : txtSearch.Text.Trim();
                    LoadItems();
                }
            }
        }

        private void LoadItems()
        {
            try
            {
                if (string.IsNullOrEmpty(ListName))
                {
                    ShowMessage("يرجى تحديد اسم القائمة من خصائص الويب بارت.", true);
                    return;
                }

                using (SPSite site = new SPSite(SPContext.Current.Site.Url))
                using (SPWeb web = site.OpenWeb())
                {
                    SPList list = web.Lists.TryGetList(ListName);
                    if (list == null)
                    {
                        ShowMessage($"القائمة '{ListName}' غير موجودة.", true);
                        return;
                    }

                    List<string> displayNamesList = new List<string>();
                    if (!string.IsNullOrEmpty(Columns))
                    {
                        displayNamesList = Columns.Split(',').Select(c => c.Trim()).ToList();
                    }
                    else
                    {
                        displayNamesList.Add("Title");
                    }

                    Dictionary<string, string> displayToInternalMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    Dictionary<string, string> internalToDisplayMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    displayToInternalMap["ID"] = "ID";
                    internalToDisplayMap["ID"] = "ID";

                    foreach (SPField field in list.Fields)
                    {
                        if (!field.Hidden)
                        {
                            displayToInternalMap[field.Title] = field.InternalName;
                            internalToDisplayMap[field.InternalName] = field.Title;
                        }
                    }

                    List<string> columnsList = new List<string>();
                    foreach (string displayName in displayNamesList)
                    {
                        if (displayToInternalMap.ContainsKey(displayName))
                        {
                            columnsList.Add(displayToInternalMap[displayName]);
                        }
                        else
                        {
                            columnsList.Add(displayName);
                        }
                    }

                    if (!columnsList.Contains("ID") && !columnsList.Contains("Id"))
                    {
                        columnsList.Add("ID");
                    }

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
                        ShowMessage("لا توجد عناصر في القائمة.", false);
                        _originalData = null;
                        _filteredData = null;
                        BindGridView();
                        return;
                    }

                    DataTable dt = new DataTable();
                    dt.Columns.Add("Id", typeof(int));

                    foreach (string columnInternalName in columnsList)
                    {
                        if (columnInternalName.Equals("ID", StringComparison.OrdinalIgnoreCase))
                            continue;

                        if (!internalToDisplayMap.ContainsKey(columnInternalName))
                        {
                            try
                            {
                                SPField field = list.Fields.GetFieldByInternalName(columnInternalName);
                                internalToDisplayMap[columnInternalName] = field == null ? columnInternalName : field.Title;
                            }
                            catch
                            {
                                internalToDisplayMap[columnInternalName] = columnInternalName;
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

                    foreach (SPListItem item in items)
                    {
                        DataRow dr = dt.NewRow();
                        dr["Id"] = item.ID;

                        foreach (string column in columnsList)
                        {
                            if (column.Equals("ID", StringComparison.OrdinalIgnoreCase))
                                continue;

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

                        dt.Rows.Add(dr);
                    }

                    _originalData = dt.Copy();

                    if (!string.IsNullOrEmpty(_searchTerm))
                    {
                        ApplySearchFilter();
                    }
                    else
                    {
                        _filteredData = _originalData;
                    }

                    BindGridView();

                    if (gvEvents.HeaderRow != null)
                    {
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
                ShowMessage($"حدث خطأ أثناء تحميل البيانات: {ex.Message}", true);
            }
        }

        protected void gvEvents_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gvEvents.PageIndex = e.NewPageIndex;

            if (_filteredData == null)
            {
                _searchTerm = txtSearch == null ? string.Empty : txtSearch.Text.Trim();
                LoadItems();
                return;
            }

            BindGridView();
        }

        protected void btnSearch_Click(object sender, EventArgs e)
        {
            _searchTerm = txtSearch.Text.Trim();
            gvEvents.PageIndex = 0;

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
                LoadItems();
            }
        }

        protected void btnGoToPage_Click(object sender, EventArgs e)
        {
            if (_filteredData == null)
            {
                _searchTerm = txtSearch == null ? string.Empty : txtSearch.Text.Trim();
                LoadItems();
            }

            if (_filteredData == null)
                return;

            int pageNumber;
            if (int.TryParse(txtPageNumber.Text, out pageNumber))
            {
                pageNumber = pageNumber - 1;

                int totalPages = (int)Math.Ceiling((double)_filteredData.Rows.Count / gvEvents.PageSize);
                if (pageNumber >= 0 && pageNumber < totalPages)
                {
                    gvEvents.PageIndex = pageNumber;
                    BindGridView();
                }
                else
                {
                    ShowMessage($"يرجى إدخال رقم صفحة صحيح بين 1 و {totalPages}", true);
                }
            }
            else
            {
                ShowMessage("يرجى إدخال رقم صفحة صحيح", true);
            }
        }

        private void ApplySearchFilter()
        {
            if (_originalData == null || string.IsNullOrEmpty(_searchTerm))
                return;

            _filteredData = _originalData.Clone();

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
            {
                gvEvents.DataSource = null;
                gvEvents.DataBind();

                lblCurrentPage.Text = "0";
                lblTotalPages.Text = "0";
                lblTotalRecords.Text = "0";
                return;
            }

            DataView dv = new DataView(_filteredData);
            gvEvents.DataSource = dv;
            gvEvents.DataBind();

            int totalRecords = _filteredData.Rows.Count;
            int totalPages = (int)Math.Ceiling((double)totalRecords / gvEvents.PageSize);

            lblCurrentPage.Text = (gvEvents.PageIndex + 1).ToString();
            lblTotalPages.Text = totalPages.ToString();
            lblTotalRecords.Text = totalRecords.ToString();

            txtPageNumber.Attributes["max"] = totalPages.ToString();
        }

        protected void gvEvents_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
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
                                e.Row.Cells[cellIndex].Text = dateValue.ToString("dd/MM/yyyy");
                            }
                        }
                    }

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
                e.Row.CssClass = "brown-bg text-white";

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
