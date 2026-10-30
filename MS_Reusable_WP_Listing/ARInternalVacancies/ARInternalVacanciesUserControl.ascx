<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ARInternalVacanciesUserControl.ascx.cs" Inherits="MS_Reusable_WP_Listing.ARInternalVacancies.ARInternalVacanciesUserControl" %>

<script src="/Style Library/CBE/JS/jquery-ui.js"></script>

<style type="text/css">
    /* RTL Styles for Arabic */
    .ar-container {
        direction: rtl;
        text-align: right;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    .search-container {
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        flex-direction: row-reverse; /* Reverse for RTL */
    }
    
    .search-input {
        flex-grow: 1;
        padding: 8px 12px;
        border: 1px solid #ddd;
        border-radius: 0 4px 4px 0; /* Reversed for RTL */
        font-size: 14px;
        text-align: right;
    }
    
    .search-button {
        background-color: #a9793f; /* Havana color */
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 4px 0 0 4px; /* Reversed for RTL */
        cursor: pointer;
        font-size: 14px;
    }
    
    .search-button:hover {
        background-color: #8a6433;
    }
    
    .search-label {
        margin-left: 10px; /* Reversed for RTL */
        font-weight: bold;
    }
    
    .pagination-container {
        display: flex;
        align-items: center;
        justify-content: center;
        margin-top: 20px;
        flex-direction: row-reverse; /* Reverse for RTL */
    }
    
    .page-jump {
        display: flex;
        align-items: center;
        margin-right: 15px; /* Reversed for RTL */
        flex-direction: row-reverse; /* Reverse for RTL */
    }
    
    .page-jump-input {
        width: 60px;
        padding: 4px 8px;
        border: 1px solid #ddd;
        border-radius: 4px;
        margin: 0 5px;
        text-align: center;
    }
    
    .page-jump-button {
        background-color: #a9793f;
        color: white;
        border: none;
        padding: 5px 10px;
        border-radius: 4px;
        cursor: pointer;
    }
    
    .page-jump-button:hover {
        background-color: #8a6433;
    }
    
    .page-info {
        margin-left: 15px; /* Reversed for RTL */
        color: #666;
    }
    
    .pdf-link {
        color: #a9793f;
        text-decoration: none;
        display: flex;
        align-items: flex-start;
        justify-content: flex-start;
    }
    
    .pdf-link i {
        margin-right: 5px;
    }
    
    .pdf-link:hover {
        text-decoration: underline;
    }
    
    /* RTL table styles */
    .table-responsive {
        direction: rtl;
    }
    
    .table th, .table td {
        text-align: right;
    }
    
    /* Fix for pagination arrows in RTL */
    .pagination > li:first-child > a {
        border-radius: 0 4px 4px 0;
    }
    
    .pagination > li:last-child > a {
        border-radius: 4px 0 0 4px;
    }
</style>

<asp:UpdatePanel ID="UpdatePanel1" runat="server">
    <ContentTemplate>
        <div class="ar-container">
            <asp:Panel ID="pnlMsg" runat="server" Visible="false">
                <div class="note-message warning-color d-flex w-100 ">
                    <i class="fa fa-exclamation-triangle" aria-hidden="true"></i>
                    <h4 class="note-title">
                        <asp:Literal runat="server" ID="ltr_Msg"></asp:Literal>
                    </h4>
                </div>
            </asp:Panel>
            
            <div class="card-container d-inline-block w-100 p-4">
                <!-- Search Box -->
                <div class="search-container">
                    <asp:Button ID="btnSearch" runat="server" Text="بحث" CssClass="search-button" OnClick="btnSearch_Click" />
                    <asp:TextBox ID="txtSearch" runat="server" CssClass="search-input" placeholder="أدخل كلمة البحث..."></asp:TextBox>
                    <span class="search-label">بحث:</span>
                </div>
                
                <div class="table-responsive">
                    <asp:GridView ID="gvVacancies" runat="server" class="table table-hover table-bordered table-striped customTable" PageSize="20" AllowPaging="true"
                        OnPageIndexChanging="gvVacancies_PageIndexChanging" OnRowDataBound="gvVacancies_RowDataBound" AutoGenerateColumns="false" DataKeyNames="Id"
                        HeaderStyle-CssClass="brown-bg text-white" RowStyle-CssClass="align-middle">
                        <PagerSettings Mode="NextPrevious" FirstPageText="&raquo;" LastPageText="&laquo;" NextPageText="&lsaquo;" PreviousPageText="&rsaquo;" />
                        <PagerStyle CssClass="pagination justify-content-center" HorizontalAlign="Center" />
                        <EmptyDataTemplate>
                            <asp:Literal ID="ltr_Empty" runat="server" Text="لم يتم العثور على وظائف شاغرة."></asp:Literal>
                        </EmptyDataTemplate>
                    </asp:GridView>
                    
                    <!-- Page Jump Section -->
                    <div class="pagination-container">
                        <div class="page-jump">
                            <asp:Button ID="btnGoToPage" runat="server" Text="اذهب" CssClass="page-jump-button" OnClick="btnGoToPage_Click" />
                            <asp:TextBox ID="txtPageNumber" runat="server" CssClass="page-jump-input" TextMode="Number" min="1"></asp:TextBox>
                            <span>انتقل إلى الصفحة:</span>
                        </div>
                        <div class="page-info">
                            الصفحة <asp:Label ID="lblCurrentPage" runat="server" Text="1"></asp:Label> من 
                            <asp:Label ID="lblTotalPages" runat="server" Text="1"></asp:Label>
                            (<asp:Label ID="lblTotalRecords" runat="server" Text="0"></asp:Label> سجل)
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </ContentTemplate>
</asp:UpdatePanel>
