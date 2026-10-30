<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ReusablePictureWPUserControl.ascx.cs" Inherits="MS_Reusable_WP_Listing.ReusablePictureWP.ReusablePictureWPUserControl" %>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css" />

<style type="text/css">
    /* Picture card styles */
    .picture-card {
        transition: all 0.3s ease;
        border: 1px solid #eee;
        margin-bottom: 20px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        overflow: hidden;
    }
    .picture-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 5px 15px rgba(0,0,0,0.2);
    }
    .card-img-top {
        height: 180px;
        object-fit: cover;
        position: relative;
        width: 100%;
        border-radius: 4px 4px 0 0;
    }
    .card-body {
        padding: 1rem;
        height: 220px; /* Fixed height for card body */
        position: relative;
    }
    .date-badge {
        color: #a9793f;
        font-size: 0.85rem;
        margin-bottom: 8px;
    }
    .card-title {
        font-size: 1.1rem;
        font-weight: 600;
        line-height: 1.3;
        margin-bottom: 0.5rem;
        color: #333;
        text-align: left;
        height: 2.8rem; /* Fixed height for title - approximately 2 lines */
        overflow: hidden;
    }
    .card-text-container {
        height: 80px; /* Fixed height for description area */
        overflow: hidden;
        margin-bottom: 15px;
    }
    .card-text {
        color: #666;
        font-size: 0.9rem;
        line-height: 1.5;
        margin-bottom: 0;
    }
    .read-more-wrapper {
        position: absolute;
        bottom: 1rem;
        left: 1rem;
        right: 1rem;
    }
    .read-more-link {
        color: #fff;
        font-size: 0.8rem;
        font-weight: 600;
        text-decoration: none;
        padding: 0;
        display: flex;
        align-items: stretch;
        justify-content: space-between;
        position: relative;
        overflow: hidden;
        height: 40px;
    }
    .read-more-link span {
        background-color: #333;
        color: #fff;
        padding: 0 15px;
        display: flex;
        align-items: center;
        flex-grow: 1;
        position: relative;
        z-index: 1;
        text-transform: uppercase;
        font-size: 0.75rem;
        letter-spacing: 0.5px;
        margin-right: -20px;
        padding-right: 35px;
    }
    .read-more-link span:after {
        content: '';
        position: absolute;
        top: 0;
        right: 0;
        width: 30px;
        height: 100%;
        background-color: #d9534f;
        transform: skewX(-37deg);
        transform-origin: top right;
        z-index: -1;
    }
    .read-more-link.gold span:after {
        background-color: #a9793f;
    }
    .read-more-link.gold span {
        background-color: #333;
        color: #a9793f;
    }
    .read-more-link i {
        background-color: #d9534f;
        color: #fff;
        width: 40px;
        display: flex;
        align-items: center;
        justify-content: center;
        position: relative;
        z-index: 2;
        margin-left: -10px;
    }
    .read-more-link.gold i {
        background-color: #a9793f;
    }
    .read-more-link:hover {
        text-decoration: none;
    }
    .pictures-header {
        position: relative;
        margin-bottom: 25px;
        border-bottom: 1px solid #eee;
        padding-bottom: 10px;
    }
    .pictures-header h2 {
        font-size: 1.5rem;
        font-weight: 600;
        color: #333;
        margin: 0;
        padding: 0;
    }
    .title-bar {
        display: inline-block;
        width: 4px;
        height: 20px;
        background-color: #a9793f;
        margin-right: 10px;
        vertical-align: middle;
    }
    
    /* Container */
    .reusable-pictures-container {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        max-width: 100%;
        margin: 0 auto;
        font-size: 14px;
        border: 1px solid #ddd;
        background-color: white;
        padding: 15px;
    }
    
    /* Row and column layout */
    .row {
        display: flex;
        flex-wrap: wrap;
        margin-right: -15px;
        margin-left: -15px;
    }
    .col-lg-4 {
        position: relative;
        width: 33.333333%;
        padding-right: 15px;
        padding-left: 15px;
    }
    
    /* Pagination */
    .pager-wrapper {
        margin-top: 20px;
    }
    .pager a {
        padding: 5px 10px;
        margin: 0 2px;
        border: 1px solid #ddd;
        color: #333;
        text-decoration: none;
        border-radius: 3px;
    }
    .pager a:hover {
        background-color: #f5f5f5;
    }
    .btn-primary.disabled {
        background-color: #a9793f;
        border-color: #a9793f;
        color: white;
    }
    
    /* Error Message */
    .error-message {
        padding: 15px;
        background-color: #f8d7da;
        color: #721c24;
        border: 1px solid #f5c6cb;
        border-radius: 4px;
        margin-bottom: 15px;
    }
    
    /* Responsive */
    @media (max-width: 992px) {
        .col-lg-4 {
            width: 50%;
        }
    }
    @media (max-width: 768px) {
        .col-lg-4 {
            width: 100%;
        }
    }
    
    /* Utility classes */
    .d-flex {
        display: flex !important;
    }
    .justify-content-center {
        justify-content: center !important;
    }
    .align-items-center {
        align-items: center !important;
    }
    .d-inline-block {
        display: inline-block !important;
    }
    .h-100 {
        height: 100% !important;
    }
    .p-3 {
        padding: 1rem !important;
    }
    .mb-2 {
        margin-bottom: 0.5rem !important;
    }
    .mb-3 {
        margin-bottom: 1rem !important;
    }
</style>

<asp:UpdatePanel ID="UpdatePanel1" runat="server">
    <ContentTemplate>
        <div class="reusable-pictures-container">
            <div class="pictures-header d-flex align-items-center">
                <div class="title-bar"></div>
                <h2></h2>
            </div>
            
            <asp:Panel ID="pnlContent" runat="server">
                <asp:ListView runat="server" ID="lsvPictures" OnPagePropertiesChanging="lsvPictures_PagePropertiesChanging">
                    <LayoutTemplate>
                        <div class="row">
                            <div runat="server" id="itemPlaceholder"></div>
                        </div>
                        <div class="pager-wrapper d-flex justify-content-center">
                            <div class="pager d-inline-block">
                                <asp:DataPager ID="_dpPager" runat="server" PagedControlID="lsvPictures" PageSize="6">
                                    <Fields>
                                        <asp:NextPreviousPagerField PreviousPageText="<" FirstPageText="|<" ShowPreviousPageButton="true"
                                            ShowFirstPageButton="true" ShowNextPageButton="false" ShowLastPageButton="false"
                                            RenderNonBreakingSpacesBetweenControls="false" RenderDisabledButtonsAsLabels="false" />
                                        <asp:NumericPagerField ButtonType="Link" CurrentPageLabelCssClass="btn btn-primary disabled" RenderNonBreakingSpacesBetweenControls="false"
                                            NextPageText="..." NextPreviousButtonCssClass="btn btn-default" />
                                        <asp:NextPreviousPagerField NextPageText=">" LastPageText=">|" ShowNextPageButton="true"
                                            ShowLastPageButton="true" ShowPreviousPageButton="false" ShowFirstPageButton="false"
                                            RenderNonBreakingSpacesBetweenControls="false" RenderDisabledButtonsAsLabels="false" />
                                    </Fields>
                                </asp:DataPager>
                            </div>
                        </div>
                    </LayoutTemplate>
                    <ItemTemplate>
                        <div class="col-lg-4">
                            <div class="picture-card h-100">
                                <img src="/PublishingImages/reuse%20wp/imereu.jpg" class="card-img-top" alt='<%# Eval("Title") %>'>

                                <div class="card-body">
                                    <div class="date-badge"><%# FormatDate(Container.DataItem) %></div>
                                    <h4 class="card-title"><%# Eval("Title") %></h4>
                                    <div class="card-text-container">
                                        <p class="card-text"><%# GetDescription(Container.DataItem) %></p>
                                    </div>
                                    <div class="read-more-wrapper">
                                        <a class="read-more-link <%# Container.DataItemIndex % 2 == 0 ? "" : "gold" %>" href='<%# GetHyperlinkInfo(Container.DataItem).Url %>'>
                                            <span><%# GetHyperlinkInfo(Container.DataItem).DisplayText %></span>
                                            <i class="fas fa-chevron-right"></i>
                                        </a>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </ItemTemplate>
                    <EmptyDataTemplate>
                        <div class="no-data">No items found in the list.</div>
                    </EmptyDataTemplate>
                </asp:ListView>
            </asp:Panel>
            
            <asp:Panel ID="pnlError" runat="server" CssClass="error-message" Visible="false">
                <asp:Literal ID="litError" runat="server"></asp:Literal>
            </asp:Panel>
        </div>
    </ContentTemplate>
</asp:UpdatePanel>