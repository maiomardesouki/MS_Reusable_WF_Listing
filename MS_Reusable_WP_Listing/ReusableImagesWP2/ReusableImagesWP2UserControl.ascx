<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ReusableImagesWP2UserControl.ascx.cs" Inherits="MS_Reusable_WP_Listing.ReusableImagesWP2.ReusableImagesWP2UserControl" %>

<style type="text/css">
    .image-card {
        transition: all 0.3s ease;
        border: 1px solid #eee;
    }
    .card-img-top {
        height: 180px;
        object-fit: cover;
        position: relative;
    }
    .image-card {
        position: relative;
    }
    .date-badge {
        color: #a9793f;
        font-size: 0.85rem;
    }
    .card-title {
        font-size: 1.1rem;
        font-weight: 600;
        line-height: 1.3;
        margin-bottom: 0.5rem;
        text-align: left !important;
        height: 2.8rem; /* Fixed height for title - approximately 2 lines */
        overflow: hidden;
        width: 100%;
        display: block;
    }
    
    .card-title span {
        display: block;
        max-height: 2.8rem;
        overflow: hidden;
        text-overflow: ellipsis;
        text-align: left !important;
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
    .images-header {
        position: relative;
        margin-bottom: 25px;
        border-bottom: 1px solid #eee;
        padding-bottom: 10px;
    }
    .images-header h2 {
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
</style>

<asp:UpdatePanel ID="UpdatePanel1" runat="server">
    <ContentTemplate>
        <div class="reusable-images-container">
            <div class="images-header d-flex align-items-center">
                <div class="title-bar"></div>
                <h2></h2>
            </div>
            
            <asp:Panel ID="pnlContent" runat="server">
                <asp:ListView runat="server" ID="lsvImages" OnPagePropertiesChanging="lsvImages_PagePropertiesChanging">
                    <LayoutTemplate>
                        <div class="row g-3">
                            <div runat="server" id="itemPlaceholder"></div>
                        </div>
                        <div class="pager-wrapper d-flex justify-content-center">
                            <div class="pager d-inline-block">
                                <asp:DataPager ID="_dpPager" runat="server" PagedControlID="lsvImages" PageSize="6">
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
                        <div class="col-lg-4 col-md-6 mb-4">
                            <div class="image-card h-100 shadow-sm bg-white rounded overflow-hidden <%# Container.DataItemIndex % 2 == 0 ? "" : "gold" %>">
                                <img src='http://sp19server:42514/en/SiteAssets/2cbe%20news.png' class="card-img-top" alt='<%# Eval("Title") %>'>
                                <div class="card-body p-3 d-flex flex-column">
                                    <div class="date-badge mb-2"><%# FormatDate(Container.DataItem) %></div>
                                    <h4 class="card-title mb-2"><span><%# Eval("Title") %></span></h4>
                                    <div class="card-text-container">
                                        <p class="card-text"><%# GetDescription(Container.DataItem) %></p>
                                    </div>
                                    <div class="read-more-wrapper mt-auto">
                                        <a class="read-more-link <%# Container.DataItemIndex % 2 == 0 ? "" : "gold" %>" href='<%# this.DetailsPageUrl + "?aID=" + Eval("ID") %>'>
                                            <span>READ MORE ...</span>
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

        <style type="text/css">
            /* Images Gallery Header */
            .images-header {
                width: 100%;
                display: flex;
                align-items: center;
                margin-bottom: 15px;
                padding: 0;
            }
            
            .title-bar {
                width: 6px;
                height: 24px;
                background-color: #a9793f; /* Havana color */
                margin-right: 10px;
                border-radius: 2px;
            }
            
            .images-header h2 {
                margin: 0;
                font-size: 16px;
                font-weight: bold;
                color: #333;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }
            
            /* Container */
            .reusable-images-container {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                max-width: 100%;
                margin: 0 auto;
                font-size: 14px;
                border: 1px solid #ddd;
                background-color: white;
                padding: 15px;
            }
            
            /* Image Items */
            .events-item {
                margin-bottom: 20px;
                transition: all 0.3s ease;
                overflow: hidden;
                height: 100%;
            }
            
            .events-item:hover {
                transform: translateY(-5px);
                box-shadow: 0 10px 20px rgba(0,0,0,0.1) !important;
            }
            
            .events-Pic {
                height: 200px;
                object-fit: cover;
            }
            
            .events-item .d-inline-block {
                padding: 15px;
            }
            
            .maindate {
                color: #777;
                font-size: 12px;
                display: block;
                margin-bottom: 8px;
            }
            
            .maindate i {
                margin-right: 5px;
                color: #c68c53;
            }
            
            .events-title {
                font-size: 18px;
                font-weight: bold;
                margin-bottom: 10px;
                color: #333;
                line-height: 1.3;
            }
            
            .maindescription {
                font-size: 14px;
                color: #555;
                margin-bottom: 15px;
                line-height: 1.4;
                max-width: 100%;
                overflow-wrap: break-word;
                word-break: normal;
            }
            
            .read-more {
                color: #c68c53; /* Orange/gold color */
                font-size: 14px;
                text-decoration: none;
                padding-top: 10px;
                border-top: 1px solid #eee;
            }
            
            .read-more:hover {
                color: #a06a35;
                text-decoration: none;
            }
            
            /* Pager */
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
                background-color: #c68c53;
                border-color: #c68c53;
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
                    float: left;
                }
            }
            
            @media (max-width: 768px) {
                .col-lg-4 {
                    width: 100%;
                    float: none;
                }
            }
            
            /* Bootstrap-like classes */
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
            
            .d-flex {
                display: flex !important;
            }
            
            .flex-column {
                flex-direction: column !important;
            }
            
            .mt-auto {
                margin-top: auto !important;
            }
            
            .justify-content-center {
                justify-content: center !important;
            }
            
            .justify-content-between {
                justify-content: space-between !important;
            }
            
            .align-items-center {
                align-items: center !important;
            }
            
            .d-inline-block {
                display: inline-block !important;
            }
            
            .w-100 {
                width: 100% !important;
            }
            
            .shadow-sm {
                box-shadow: 0 .125rem .25rem rgba(0,0,0,.075) !important;
            }
            
            .bg-white {
                background-color: #fff !important;
            }
            
            .rounded {
                border-radius: .25rem !important;
            }
            
            .margB10 {
                margin-bottom: 10px !important;
            }
        </style>
    </ContentTemplate>
</asp:UpdatePanel>
