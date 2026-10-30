using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System;
using System.ComponentModel;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

namespace MS_Reusable_WP_Listing.EventDetailsWP
{
    [ToolboxItemAttribute(false)]
    public class EventDetailsWP : WebPart
    {
        // Visual Studio might automatically update this path when you change the Visual Web Part project item.
        private const string _ascxPath = @"~/_CONTROLTEMPLATES/15/MS_Reusable_WP_Listing/EventDetailsWP/EventDetailsWPUserControl.ascx";

        #region Properties
        
        // Property to store the list name
        private string _listName = string.Empty;
        [WebBrowsable(true)]
        [WebDisplayName("List Name")]
        [WebDescription("Name of the SharePoint list containing events")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string ListName
        {
            get { return _listName; }
            set { _listName = value; }
        }
        
        // Property to store the title field name
        private string _titleField = "Title";
        [WebBrowsable(true)]
        [WebDisplayName("Title Field")]
        [WebDescription("Name of the field containing the event title")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Field Mapping")]
        public string TitleField
        {
            get { return _titleField; }
            set { _titleField = value; }
        }
        
        // Property to store the Arabic title field name
        private string _titleArField = "TitleAr";
        [WebBrowsable(true)]
        [WebDisplayName("Arabic Title Field")]
        [WebDescription("Name of the field containing the Arabic event title")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Field Mapping")]
        public string TitleArField
        {
            get { return _titleArField; }
            set { _titleArField = value; }
        }
        
        // Property to store the description field name
        private string _descriptionField = "Description";
        [WebBrowsable(true)]
        [WebDisplayName("Description Field")]
        [WebDescription("Name of the field containing the event description")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Field Mapping")]
        public string DescriptionField
        {
            get { return _descriptionField; }
            set { _descriptionField = value; }
        }
        
        // Property to store the Arabic description field name
        private string _descriptionArField = "DescriptionAr";
        [WebBrowsable(true)]
        [WebDisplayName("Arabic Description Field")]
        [WebDescription("Name of the field containing the Arabic event description")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Field Mapping")]
        public string DescriptionArField
        {
            get { return _descriptionArField; }
            set { _descriptionArField = value; }
        }
        
        // Property to store the location field name
        private string _locationField = "Location";
        [WebBrowsable(true)]
        [WebDisplayName("Location Field")]
        [WebDescription("Name of the field containing the event location")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Field Mapping")]
        public string LocationField
        {
            get { return _locationField; }
            set { _locationField = value; }
        }
        
        // Property to store the Arabic location field name
        private string _locationArField = "LocationAr";
        [WebBrowsable(true)]
        [WebDisplayName("Arabic Location Field")]
        [WebDescription("Name of the field containing the Arabic event location")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Field Mapping")]
        public string LocationArField
        {
            get { return _locationArField; }
            set { _locationArField = value; }
        }
        
        // Property to store the start date field name
        private string _startDateField = "StartDate";
        [WebBrowsable(true)]
        [WebDisplayName("Start Date Field")]
        [WebDescription("Name of the field containing the event start date")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Field Mapping")]
        public string StartDateField
        {
            get { return _startDateField; }
            set { _startDateField = value; }
        }
        
        // Property to store the end date field name
        private string _endDateField = "EndDate";
        [WebBrowsable(true)]
        [WebDisplayName("End Date Field")]
        [WebDescription("Name of the field containing the event end date")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Field Mapping")]
        public string EndDateField
        {
            get { return _endDateField; }
            set { _endDateField = value; }
        }
        
        // Property to store the category field name
        private string _categoryField = "CalendarCategory";
        [WebBrowsable(true)]
        [WebDisplayName("Category Field")]
        [WebDescription("Name of the field containing the event category")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Field Mapping")]
        public string CategoryField
        {
            get { return _categoryField; }
            set { _categoryField = value; }
        }
        
        // Property to store the approval field name
        private string _approvedField = "IsApproved";
        [WebBrowsable(true)]
        [WebDisplayName("Approval Field")]
        [WebDescription("Name of the Yes/No field indicating if the event is approved")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Field Mapping")]
        public string ApprovedField
        {
            get { return _approvedField; }
            set { _approvedField = value; }
        }
        
        // Property to determine if we should use Arabic fields
        private bool _useArabicFields = false;
        [WebBrowsable(true)]
        [WebDisplayName("Use Arabic Fields")]
        [WebDescription("Set to true to display Arabic content when available")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public bool UseArabicFields
        {
            get { return _useArabicFields; }
            set { _useArabicFields = value; }
        }
        
        #endregion

        protected override void CreateChildControls()
        {
            Control control = Page.LoadControl(_ascxPath);
            
            // Pass all properties to the user control
            if (control is EventDetailsWPUserControl)
            {
                var userControl = (EventDetailsWPUserControl)control;
                userControl.ListName = this.ListName;
                userControl.TitleField = this.TitleField;
                userControl.TitleArField = this.TitleArField;
                userControl.DescriptionField = this.DescriptionField;
                userControl.DescriptionArField = this.DescriptionArField;
                userControl.LocationField = this.LocationField;
                userControl.LocationArField = this.LocationArField;
                userControl.StartDateField = this.StartDateField;
                userControl.EndDateField = this.EndDateField;
                userControl.CategoryField = this.CategoryField;
                userControl.ApprovedField = this.ApprovedField;
                userControl.UseArabicFields = this.UseArabicFields;
            }
            
            Controls.Add(control);
        }
    }
}
