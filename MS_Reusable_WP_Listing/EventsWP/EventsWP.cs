using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System;
using System.ComponentModel;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

namespace MS_Reusable_WP_Listing.EventsWP
{
    [ToolboxItemAttribute(false)]
    public class EventsWP : WebPart
    {
        // Visual Studio might automatically update this path when you change the Visual Web Part project item.
        private const string _ascxPath = @"~/_CONTROLTEMPLATES/15/MS_Reusable_WP_Listing/EventsWP/EventsWPUserControl.ascx";
        // Define the ListName property to be exposed in the web part properties pane
        private string _listName = string.Empty;

        [WebBrowsable(true)]
        [WebDisplayName("List Name")]
        [WebDescription("Enter the name of the SharePoint calendar list to display")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string ListName
        {
            get { return _listName; }
            set { _listName = value; }
        }

        // Define the CalendarTitle property
        private string _calendarTitle = "Calendar";

        [WebBrowsable(true)]
        [WebDisplayName("Calendar Title")]
        [WebDescription("Enter the title to display for the calendar")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string CalendarTitle
        {
            get { return _calendarTitle; }
            set { _calendarTitle = value; }
        }

        protected override void CreateChildControls()
        {
            Control control = Page.LoadControl(_ascxPath);

            // Pass the properties to the user control
            if (control is EventsWPUserControl)
            {
                ((EventsWPUserControl)control).ListName = this.ListName;
                ((EventsWPUserControl)control).CalendarTitle = this.CalendarTitle;
            }

            Controls.Add(control);
        }
    }
}
