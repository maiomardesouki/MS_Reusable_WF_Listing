using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System;
using System.ComponentModel;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

namespace MS_Reusable_WP_Listing.AR_EmployeeHandbook
{
    [ToolboxItemAttribute(false)]
    public class AR_EmployeeHandbook : WebPart
    {
        // Visual Studio might automatically update this path when you change the Visual Web Part project item.
        private const string _ascxPath = @"~/_CONTROLTEMPLATES/15/MS_Reusable_WP_Listing/AR_EmployeeHandbook/AR_EmployeeHandbookUserControl.ascx";

        private string _listName = string.Empty;

        [WebBrowsable(true)]
        [WebDisplayName("List Name")]
        [WebDescription("Enter the name of the SharePoint list to display")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string ListName
        {
            get { return _listName; }
            set { _listName = value; }
        }

        private string _columns = string.Empty;

        [WebBrowsable(true)]
        [WebDisplayName("Columns")]
        [WebDescription("Enter the display names of columns to show, separated by commas. You can use either display names or internal names.")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string Columns
        {
            get { return _columns; }
            set { _columns = value; }
        }

        protected override void CreateChildControls()
        {
            Control control = Page.LoadControl(_ascxPath);

            if (control is AR_EmployeeHandbookUserControl)
            {
                ((AR_EmployeeHandbookUserControl)control).ListName = this.ListName;
                ((AR_EmployeeHandbookUserControl)control).Columns = this.Columns;
            }

            Controls.Add(control);
        }
    }
}
