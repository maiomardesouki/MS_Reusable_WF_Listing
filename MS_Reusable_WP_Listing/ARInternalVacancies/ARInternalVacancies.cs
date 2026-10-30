using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System;
using System.ComponentModel;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

namespace MS_Reusable_WP_Listing.ARInternalVacancies
{
    [ToolboxItemAttribute(false)]
    public class ARInternalVacancies : WebPart
    {
        // Visual Studio might automatically update this path when you change the Visual Web Part project item.
        private const string _ascxPath = @"~/_CONTROLTEMPLATES/15/MS_Reusable_WP_Listing/ARInternalVacancies/ARInternalVacanciesUserControl.ascx";

        // Web Part Properties
        private string _listName = "Internal Vacancies";
        [WebBrowsable(true)]
        [WebDisplayName("List Name")]
        [WebDescription("The name of the SharePoint list containing the vacancies data")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string ListName
        {
            get { return _listName; }
            set { _listName = value; }
        }

        private string _columns = "Title,Department,Location,ClosingDate";
        [WebBrowsable(true)]
        [WebDisplayName("Columns")]
        [WebDescription("Comma-separated list of column internal names to show in the table")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string Columns
        {
            get { return _columns; }
            set { _columns = value; }
        }
        
        private string _columnDisplayNames = "Title:عنوان الوظيفة,Department:القسم,Location:الموقع,ClosingDate:تاريخ الإغلاق";
        [WebBrowsable(true)]
        [WebDisplayName("Column Display Names")]
        [WebDescription("Comma-separated list of column display names in format 'InternalName:ArabicDisplayName'")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string ColumnDisplayNames
        {
            get { return _columnDisplayNames; }
            set { _columnDisplayNames = value; }
        }

        private string _pdfField = "JobDescription";
        [WebBrowsable(true)]
        [WebDisplayName("PDF Field")]
        [WebDescription("The field name containing the PDF URL")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string PdfField
        {
            get { return _pdfField; }
            set { _pdfField = value; }
        }

        protected override void CreateChildControls()
        {
            Control control = Page.LoadControl(_ascxPath);
            
            // Pass properties to the user control
            if (control is ARInternalVacanciesUserControl)
            {
                ARInternalVacanciesUserControl userControl = (ARInternalVacanciesUserControl)control;
                userControl.ListName = this.ListName;
                userControl.Columns = this.Columns;
                userControl.ColumnDisplayNames = this.ColumnDisplayNames;
                userControl.PdfField = this.PdfField;
            }
            
            Controls.Add(control);
        }
    }
}
