<%@ Assembly Name="MS_Reusable_WP_Listing, Version=1.0.0.0, Culture=neutral, PublicKeyToken=496cdbb56b5c3083" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="InternalVacanciesUserControl.ascx.cs" Inherits="MS_Reusable_WP_Listing.InternalVacancies.InternalVacanciesUserControl" %>

<script src="/Style Library/CBE/JS/jquery-ui.js"></script>

<style type="text/css">
    .search-container {
        margin-bottom: 20px;
        display: flex;
        align-items: center;
    }
    
    .search-input {
        flex-grow: 1;
        padding: 8px 12px;
        border: 1px solid #ddd;
        border-radius: 4px 0 0 4px;
        font-size: 14px;
    }
    
    .search-button {
        background-color: #a9793f; /* Havana color */
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 0 4px 4px 0;
        cursor: pointer;
        font-size: 14px;
    }
    
    .search-button:hover {
        background-color: #8a6433;
    }
    
    .search-label {
        margin-right: 10px;
        font-weight: bold;
    }
    
    .pagination-container {
        display: flex;
        align-items: center;
        justify-content: center;
        margin-top: 20px;
    }
    
    .page-jump {
        display: flex;
        align-items: center;
        margin-left: 15px;
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
        margin-right: 15px;
        color: #666;
    }
    
    .pdf-link {
        color: #a9793f;
        text-decoration: none;
        display: flex;
        align-items: left;
        justify-content: left;
    }
    
    .pdf-link i {
        margin-right: 5px;
    }
    
    .pdf-link:hover {
        text-decoration: underline;
    }
</style>

<asp:UpdatePanel ID="UpdatePanel1" runat="server">
    <ContentTemplate>
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
                <span class="search-label">Search:</span>
                <asp:TextBox ID="txtSearch" runat="server" CssClass="search-input" placeholder="Enter search term..."></asp:TextBox>
                <asp:Button ID="btnSearch" runat="server" Text="Search" CssClass="search-button" OnClick="btnSearch_Click" />
            </div>
            
            <div class="table-responsive">
                <asp:GridView ID="gvVacancies" runat="server" class="table table-hover table-bordered table-striped customTable" PageSize="20" AllowPaging="true"
                    OnPageIndexChanging="gvVacancies_PageIndexChanging" OnRowDataBound="gvVacancies_RowDataBound" AutoGenerateColumns="false" DataKeyNames="Id"
                    HeaderStyle-CssClass="brown-bg text-white" RowStyle-CssClass="align-middle">
                    <PagerSettings Mode="NextPrevious" FirstPageText="&laquo;" LastPageText="&raquo;" NextPageText="&rsaquo;" PreviousPageText="&lsaquo;" />
                    <PagerStyle CssClass="pagination justify-content-center" HorizontalAlign="Center" />
                    <EmptyDataTemplate>
                        <asp:Literal ID="ltr_Empty" runat="server" Text="No vacancies found."></asp:Literal>
                    </EmptyDataTemplate>
                </asp:GridView>
                
                <!-- Page Jump Section -->
                <div class="pagination-container">
                    <div class="page-info">
                        Page <asp:Label ID="lblCurrentPage" runat="server" Text="1"></asp:Label> of 
                        <asp:Label ID="lblTotalPages" runat="server" Text="1"></asp:Label>
                        (<asp:Label ID="lblTotalRecords" runat="server" Text="0"></asp:Label> records)
                    </div>
                    <div class="page-jump">
                        <span>Go to page:</span>
                        <asp:TextBox ID="txtPageNumber" runat="server" CssClass="page-jump-input" TextMode="Number" min="1"></asp:TextBox>
                        <asp:Button ID="btnGoToPage" runat="server" Text="Go" CssClass="page-jump-button" OnClick="btnGoToPage_Click" />
                    </div>
                </div>
            </div>
        </div>
    </ContentTemplate>
</asp:UpdatePanel>
