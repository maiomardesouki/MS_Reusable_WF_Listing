<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="EventDetailsWPUserControl.ascx.cs" Inherits="MS_Reusable_WP_Listing.EventDetailsWP.EventDetailsWPUserControl" %>

<div class="event-detail-container">
    <div class="event-detail-card">
        <div class="event-header">
            <!-- Yellow/gold accent bar similar to the screenshot -->
            <div class="accent-bar"></div>
            <!-- Title with proper styling -->
            <h2 class="event-title">
                <asp:Literal ID="litTitle" runat="server" />
            </h2>
            <!-- Date with calendar icon -->
            <div class="event-date">
                <i class="far fa-calendar"></i>
                <asp:Literal ID="litDate" runat="server" />
            </div>
            <!-- Hidden category for pending approval tag only -->
            <div class="event-category" style="display: none;">
                <asp:Literal ID="litCategory" runat="server" />
            </div>
        </div>
        
        <!-- Event description with proper styling for multiline text -->
        <div class="event-description">
            <asp:Literal ID="litDescription" runat="server" />
        </div>
        
        <!-- External link if available -->
        <div class="external-link" runat="server" id="divExternalLink" visible="false">
            <asp:Literal ID="litExternalLink" runat="server" />
        </div>
        
        <!-- Images section -->
        <div class="event-images" runat="server" id="divImages" visible="false">
            <asp:Literal runat="server" ID="litImages"></asp:Literal>
        </div>
        
        <!-- Links section 
        <div class="event-links" runat="server" id="divLinks" visible="false">
            <h3>External Links</h3>
            <ul>
                <asp:Literal runat="server" ID="litLinks"></asp:Literal>
            </ul>
        </div>-->

        <!-- Attachments section 
        <div class="event-attachments" runat="server" id="divAttachments" visible="false">
            <h3>Attachments</h3>
            <ul>
                <asp:Literal runat="server" ID="litAttachments"></asp:Literal>
            </ul>
        </div>-->
    </div>
</div>

<!-- CSS styles to match the screenshot layout -->
<style type="text/css">
    .event-detail-container {
        font-family: 'Segoe UI', Arial, sans-serif;
        padding: 20px;
        max-width: 100%;
    }
    
    .event-detail-card {
        background-color: #ffffff;
        border-radius: 4px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        padding: 20px;
        position: relative;
    }
    
    .event-header {
        position: relative;
        padding-left: 15px;
        margin-bottom: 20px;
    }
    
    .accent-bar {
        position: absolute;
        left: 0;
        top: 0;
        width: 6px;
        height: 100%;
        background-color: #a9793f; /* Gold/brown color from screenshot */
        border-radius: 3px;
    }
    
    .event-title {
        font-size: 24px;
        font-weight: bold;
        color: #333333;
        margin-top: 0;
        margin-bottom: 10px;
        line-height: 1.3;
    }
    
    .event-date {
        color: #666666;
        font-size: 14px;
        margin-bottom: 10px;
    }
    
    .event-date i {
        margin-right: 5px;
        color: #a9793f; /* Gold/brown color from screenshot */
    }
    
    .event-category {
        margin-top: 10px;
    }
    
    .category-tag {
        display: inline-block;
        background-color: #f0f0f0;
        padding: 4px 10px;
        border-radius: 15px;
        font-size: 12px;
        color: #666;
        margin-right: 5px;
    }
    
    .pending-tag {
        display: inline-block;
        background-color: #fff3cd;
        color: #856404;
        padding: 4px 10px;
        border-radius: 15px;
        font-size: 12px;
        margin-left: 10px;
        vertical-align: middle;
    }
    
    .event-description {
        margin: 20px 0;
        line-height: 1.6;
        color: #333333;
        font-size: 16px;
    }
    
    .event-description p {
        margin-bottom: 15px;
    }
    
    .event-links, .event-attachments {
        margin-top: 20px;
        padding-top: 15px;
        border-top: 1px solid #eee;
    }
    
    .event-links h3, .event-attachments h3 {
        font-size: 18px;
        margin-bottom: 10px;
        color: #333;
    }
    
    .event-links ul, .event-attachments ul {
        list-style: none;
        padding: 0;
        margin: 0;
    }
    
    .event-links li, .event-attachments li {
        margin-bottom: 8px;
    }
    
    .event-links a, .event-attachments a {
        color: #a9793f; /* Gold/brown color from screenshot */
        text-decoration: none;
    }
    
    .event-links a:hover, .event-attachments a:hover {
        text-decoration: underline;
    }
    
    .event-images {
        margin: 20px 0;
    }
    
    .event-image {
        max-width: 100%;
        height: auto;
        border-radius: 4px;
        margin-bottom: 15px;
    }
    
    .external-link {
        margin-top: 15px;
    }
    
    .btn-external {
        display: inline-block;
        padding: 8px 15px;
        background-color: #a9793f; /* Gold/brown color from screenshot */
        color: white;
        text-decoration: none;
        border-radius: 4px;
        font-size: 14px;
    }
    
    .btn-external:hover {
        background-color: #8a6433;
    }
    
    .event-location {
        display: block;
        margin-top: 5px;
        color: #666;
    }
    
    .event-location i {
        margin-right: 5px;
    }
</style>
