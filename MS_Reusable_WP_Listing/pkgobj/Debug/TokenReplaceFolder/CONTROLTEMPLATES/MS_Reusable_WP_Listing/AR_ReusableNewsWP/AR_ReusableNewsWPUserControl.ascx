<%@ Assembly Name="MS_Reusable_WP_Listing, Version=1.0.0.0, Culture=neutral, PublicKeyToken=496cdbb56b5c3083" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="AR_ReusableNewsWPUserControl.ascx.cs" Inherits="MS_Reusable_WP_Listing.AR_ReusableNewsWP.AR_ReusableNewsWPUserControl" %>

<style type="text/css">
    /* Banner styles */
    .news-banner {
        position: relative;
        width: 100%;
        height: 200px;
        background-image: url('/_layouts/15/images/MS_Reusable_WP_Listing/news-banner.jpg');
        background-size: cover;
        background-position: center;
        display: flex;
        align-items: center;
        justify-content: flex-end;
        padding: 0 50px;
        color: white;
        margin-bottom: 30px;
        direction: rtl;
    }
    
    .news-banner::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-color: rgba(0, 0, 0, 0.5);
    }
    
    .news-banner-content {
        position: relative;
        z-index: 1;
        text-align: right;
    }
    
    .news-banner h1 {
        font-size: 36px;
        font-weight: bold;
        margin: 0;
    }
    
    .news-banner .breadcrumb {
        display: flex;
        list-style: none;
        padding: 0;
        margin: 10px 0 0 0;
        flex-direction: row-reverse;
    }
    
    .news-banner .breadcrumb li {
        margin-left: 5px;
    }
    
    .news-banner .breadcrumb li:after {
        content: '|';
        margin-right: 5px;
    }
    
    .news-banner .breadcrumb li:last-child:after {
        content: '';
    }
    
    .news-banner .breadcrumb a {
        color: white;
        text-decoration: none;
    }
    
    /* News item styles */
    .news-container {
        max-width: 1200px;
        margin: 0 auto;
        padding: 0 15px;
        overflow: visible;
        direction: rtl;
    }
    
    .news-list {
        display: flex;
        flex-direction: column;
        gap: 0;
        overflow: visible;
        width: calc(100% - 35px);
    }
    
    .news-item-wrapper {
        position: relative;
        margin-bottom: 40px;
        padding-bottom: 40px;
        border-bottom: 1px solid #e0e0e0;
        position: relative;
    }
    
    .news-item-wrapper::after {
        content: '';
        position: absolute;
        bottom: -1px;
        right: 0;
        width: 100px;
        height: 2px;
        background-color: #e0e0e0;
    }
    
    .news-item-wrapper:last-child {
        border-bottom: none;
        margin-bottom: 0;
        padding-bottom: 0;
    }
    
    .news-item-wrapper:last-child::after {
        display: none;
    }
    
    .news-item {
        display: flex;
        padding: 20px 20px 30px 20px;
        background-color: #f9f9f9;
        border-radius: 4px;
        margin-left: 35px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        flex-direction: row-reverse;
        text-align: right;
    }
    
    .news-item-image {
        width: 200px;
        height: 150px;
        flex-shrink: 0;
        margin-left: 40px;
        position: relative;
        overflow: visible;
    }
    
    .news-item-image img {
        width: 100%;
        height: 100%;
        object-fit: cover;
    }
    
    .date-box {
        position: absolute;
        top: 70%;
        left: 0px;
        transform: translateY(-50%);
        width: 70px;
        height: 70px;
        padding: 0;
        background-color: #fff;
        text-align: center;
        z-index: 10;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        box-shadow: 0 3px 6px rgba(0,0,0,0.16);
        border: 1px solid #ddd;
    }
    
    .date-day {
        font-size: 32px;
        font-weight: bold;
        line-height: 1;
        color: #a9793f;
        display: block;
        margin-bottom: 0;
    }
    
    .date-month {
        font-size: 10px;
        font-weight: 500;
        color: #a9793f;
        display: block;
        margin-top: 0;
        line-height: 1;
    }
    
    .news-item-content {
        flex-grow: 1;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    
    /* Responsive styles */
    @media (max-width: 768px) {
        .news-item {
            flex-direction: column;
            padding: 15px 15px 25px 15px;
        }
        
        .news-item-image {
            width: 100%;
            height: 200px;
            margin-left: 0;
            margin-bottom: 20px;
        }
        
        .date-box {
            width: 80px;
            height: 80px;
            left: -20px;
            top: 50%;
        }
        
        .date-day {
            font-size: 36px;
        }
        
        .date-month {
            font-size: 12px;
        }
        
        .news-item-meta::after {
            width: 40px;
        }
    }
    
    .news-item-title {
        font-size: 20px;
        font-weight: bold;
        margin: 0 0 10px 0;
        color: #333;
    }
    
    .news-item-meta {
        display: flex;
        margin-bottom: 15px;
        color: #666;
        font-size: 14px;
        position: relative;
        flex-direction: row-reverse;
        justify-content: flex-end;
    }
    
    .news-item-meta::after {
        content: '';
        position: absolute;
        bottom: -8px;
        right: 0;
        width: 50px;
        height: 3px;
        background-color: #e74c3c;
    }
    
    .news-item-time {
        margin-left: 15px;
    }
    
    .news-item-location {
        margin-right: 15px;
    }
    
    .news-item-description {
        color: #666;
        margin-bottom: 15px;
        line-height: 1.5;
    }
    
    .read-more {
        color: #a9793f;
        text-decoration: none;
        font-weight: bold;
        font-size: 14px;
    }
    
    .read-more:hover {
        text-decoration: underline;
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
    
    /* Error message */
    .error-message {
        padding: 15px;
        background-color: #f8d7da;
        color: #721c24;
        border: 1px solid #f5c6cb;
        border-radius: 4px;
        margin-bottom: 20px;
        direction: rtl;
        text-align: right;
    }
    
    .info-message {
        padding: 15px;
        background-color: #d1ecf1;
        color: #0c5460;
        border: 1px solid #bee5eb;
        border-radius: 4px;
        margin-bottom: 20px;
        direction: rtl;
        text-align: right;
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
</style>

<asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
    <ContentTemplate>
        <!-- Banner Section -->
        
        
        <div class="news-container">
            <asp:Panel ID="pnlContent" runat="server">
                <asp:ListView ID="lsvNews" runat="server" OnPagePropertiesChanging="lsvNews_PagePropertiesChanging">
                    <LayoutTemplate>
                        <div class="news-list">
                            <div id="itemPlaceholder" runat="server"></div>
                        </div>
                        <div class="pager-wrapper d-flex justify-content-center">
                            <div class="pager d-inline-block">
                                <asp:DataPager ID="_dpPager" runat="server" PagedControlID="lsvNews" PageSize="6">
                                    <Fields>
                                        <asp:NextPreviousPagerField PreviousPageText=">" FirstPageText=">|" ShowPreviousPageButton="true"
                                            ShowFirstPageButton="true" ShowNextPageButton="false" ShowLastPageButton="false"
                                            RenderNonBreakingSpacesBetweenControls="false" RenderDisabledButtonsAsLabels="false" />
                                        <asp:NumericPagerField ButtonType="Link" CurrentPageLabelCssClass="btn btn-primary disabled" RenderNonBreakingSpacesBetweenControls="false"
                                            NextPageText="..." NextPreviousButtonCssClass="btn btn-default" />
                                        <asp:NextPreviousPagerField NextPageText="<" LastPageText="|<" ShowNextPageButton="true"
                                            ShowLastPageButton="true" ShowPreviousPageButton="false" ShowFirstPageButton="false"
                                            RenderNonBreakingSpacesBetweenControls="false" RenderDisabledButtonsAsLabels="false" />
                                    </Fields>
                                </asp:DataPager>
                            </div>
                        </div>
                    </LayoutTemplate>
                    <ItemTemplate>
                        <div class="news-item-wrapper">
                            <div class="news-item">
                                <div class="news-item-image">
                                    <img src='<%# Eval("ImageUrl") %>' alt='<%# Eval("Title") %>' />

                                    <div class="date-box">
                                        <span class="date-day"><%# Eval("DateDay") %></span>
                                        <span class="date-month"><%# Eval("DateMonth") %></span>
                                    </div>

                                </div>
                                 
                                <div class="news-item-content">
                                    <div>
                                        <h3 class="news-item-title"><%# Eval("Title") %></h3>
                                        <div class="news-item-meta">
                                            <span class="news-item-time">7:00 - 13:00</span>
                                            <span class="news-item-location"><%# Eval("Location") %></span>
                                        </div>
                                        <div class="news-item-description"><%# Eval("Description") %></div>
                                    </div>
                                    <div>
                                        <a href='<%# Eval("DetailsUrl") %>' class="read-more">اقرأ المزيد...</a>
                                    </div>
                                </div>
                            </div>
                             
                        </div>
                    </ItemTemplate>
                    <EmptyDataTemplate>
                        <div class="no-data" style="direction: rtl; text-align: right;">لا توجد أخبار.</div>
                    </EmptyDataTemplate>
                </asp:ListView>
            </asp:Panel>
            
            <asp:Panel ID="pnlError" runat="server" CssClass="error-message" Visible="false">
                <asp:Literal ID="litError" runat="server"></asp:Literal>
            </asp:Panel>
        </div>
    </ContentTemplate>
</asp:UpdatePanel>
