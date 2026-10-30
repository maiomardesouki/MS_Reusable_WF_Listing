using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System;
using System.ComponentModel;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

namespace MS_Reusable_WP_Listing.VisualWebPart1
{
    [ToolboxItemAttribute(false)]
    public class VisualWebPart1 : WebPart
    {
        // Visual Studio might automatically update this path when you change the Visual Web Part project item.
        private const string _ascxPath = @"~/_CONTROLTEMPLATES/15/MS_Reusable_WP_Listing/VisualWebPart1/VisualWebPart1UserControl.ascx";

        // Define the ListName property to be exposed in the web part properties pane
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

        // Define the Columns property to be exposed in the web part properties pane
        private string _columns = string.Empty;
        
        [WebBrowsable(true)]
        [WebDisplayName("Columns")]
        [WebDescription("Enter the display names of columns to show, separated by commas (e.g., Title,Created Date,Modified By). You can use either display names or internal names.")]
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
            
            // Pass the properties to the user control
            if (control is VisualWebPart1UserControl)
            {
                ((VisualWebPart1UserControl)control).ListName = this.ListName;
                ((VisualWebPart1UserControl)control).Columns = this.Columns;
            }
            
            Controls.Add(control);
        }
    }
}
