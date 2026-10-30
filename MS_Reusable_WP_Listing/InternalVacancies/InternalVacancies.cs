using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System;
using System.ComponentModel;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

namespace MS_Reusable_WP_Listing.InternalVacancies
{
    [ToolboxItemAttribute(false)]
    public class InternalVacancies : WebPart
    {
        // Visual Studio might automatically update this path when you change the Visual Web Part project item.
        private const string _ascxPath = @"~/_CONTROLTEMPLATES/15/MS_Reusable_WP_Listing/InternalVacancies/InternalVacanciesUserControl.ascx";

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

        private string _columns = "SECTOR,JOB_x0020_NAME,NUMBER_x0020_OF_x0020_VACANCIES,JOB_x0020_CODE,JOB_x0020_DESCRIPTION_x0020_AND_x0020_REQUIREMENTS";
        [WebBrowsable(true)]
        [WebDisplayName("Columns")]
        [WebDescription("Comma-separated list of column display names to show in the table")]
        [Personalizable(PersonalizationScope.Shared)]
        [Category("Configuration")]
        public string Columns
        {
            get { return _columns; }
            set { _columns = value; }
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
            if (control is InternalVacanciesUserControl)
            {
                InternalVacanciesUserControl userControl = (InternalVacanciesUserControl)control;
                userControl.ListName = this.ListName;
                userControl.Columns = this.Columns;
                userControl.PdfField = this.PdfField;
            }
            
            Controls.Add(control);
        }
    }
}
