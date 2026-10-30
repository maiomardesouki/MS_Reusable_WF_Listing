using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System;
using System.ComponentModel;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

namespace MS_Reusable_WP_Listing.ReusableImagesWP2
{
    [ToolboxItemAttribute(false)]
    public class ReusableImagesWP2 : WebPart
    {
        // Visual Studio might automatically update this path when you change the Visual Web Part project item.
        private const string _ascxPath = @"~/_CONTROLTEMPLATES/15/MS_Reusable_WP_Listing/ReusableImagesWP2/ReusableImagesWP2UserControl.ascx";

        // Properties for web part configuration
        private string _listName = string.Empty;
        [WebBrowsable(true)]
        [WebDisplayName("List Name")]
        [WebDescription("Enter the name of the SharePoint list to display images from")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string ListName
        {
            get { return _listName; }
            set { _listName = value; }
        }
        
        private string _webUrl = string.Empty;
        [WebBrowsable(true)]
        [WebDisplayName("Web URL")]
        [WebDescription("Enter the relative URL of the web containing the list (e.g., /en). Leave empty to use the current web.")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string WebUrl
        {
            get { return _webUrl; }
            set { _webUrl = value; }
        }

        private string _titleField = "Title";
        [WebBrowsable(true)]
        [WebDisplayName("Title Field")]
        [WebDescription("Enter the name of the field to use for the title")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string TitleField
        {
            get { return _titleField; }
            set { _titleField = value; }
        }

        private string _descriptionField = "Description";
        [WebBrowsable(true)]
        [WebDisplayName("Description Field")]
        [WebDescription("Enter the name of the field to use for the description")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string DescriptionField
        {
            get { return _descriptionField; }
            set { _descriptionField = value; }
        }

        private string _imageField = "Picture";
        [WebBrowsable(true)]
        [WebDisplayName("Image Field")]
        [WebDescription("Enter the name of the field to use for the image")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string ImageField
        {
            get { return _imageField; }
            set { _imageField = value; }
        }

        private string _dateField = "Date";
        [WebBrowsable(true)]
        [WebDisplayName("Date Field")]
        [WebDescription("Enter the name of the field to use for the date")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string DateField
        {
            get { return _dateField; }
            set { _dateField = value; }
        }

        private string _detailsPageUrl = "DetailsPage.aspx";
        [WebBrowsable(true)]
        [WebDisplayName("Details Page URL")]
        [WebDescription("Enter the URL of the details page")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string DetailsPageUrl
        {
            get { return _detailsPageUrl; }
            set { _detailsPageUrl = value; }
        }

        private int _pageSize = 6;
        [WebBrowsable(true)]
        [WebDisplayName("Page Size")]
        [WebDescription("Number of items to display per page")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public int PageSize
        {
            get { return _pageSize; }
            set { _pageSize = value; }
        }

        protected override void CreateChildControls()
        {
            Control control = Page.LoadControl(_ascxPath);
            ReusableImagesWP2UserControl userControl = (ReusableImagesWP2UserControl)control;
            
            // Pass properties to the user control
            userControl.ListName = this.ListName;
            userControl.WebUrl = this.WebUrl;
            userControl.TitleField = this.TitleField;
            userControl.DescriptionField = this.DescriptionField;
            userControl.ImageField = this.ImageField;
            userControl.DateField = this.DateField;
            userControl.DetailsPageUrl = this.DetailsPageUrl;
            userControl.PageSize = this.PageSize;
            
            Controls.Add(control);
        }
    }
}
