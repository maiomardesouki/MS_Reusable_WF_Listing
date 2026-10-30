<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="CalendarWebPartUserControl.ascx.cs" Inherits="MS_Reusable_WP_Listing.CalendarWebPart.CalendarWebPartUserControl" %>

<asp:UpdatePanel ID="UpdatePanel1" runat="server">
    <ContentTemplate>
        <div class="monthly-calendar-container">
            <div class="calendar-header">
                <h2>
                    <asp:Literal ID="ltrCalendarTitle" runat="server" Text="Calendar" />
                </h2>
                <div class="calendar-navigation">
                    <asp:LinkButton ID="btnPrevMonth" runat="server" CssClass="nav-button prev-month" OnClick="btnPrevMonth_Click">
                        <i class="fas fa-chevron-left"></i>
                    </asp:LinkButton>
                    <asp:LinkButton ID="btnToday" runat="server" CssClass="nav-button today" OnClick="btnToday_Click">Today</asp:LinkButton>
                    <asp:LinkButton ID="btnNextMonth" runat="server" CssClass="nav-button next-month" OnClick="btnNextMonth_Click">
                        <i class="fas fa-chevron-right"></i>
                    </asp:LinkButton>
                </div>
                <div class="month-display">
                    <asp:Literal ID="ltrCurrentMonth" runat="server" />
                    <asp:Literal ID="ltrCurrentYear" runat="server" />
                </div>
            </div>

            <div class="calendar-body">
                <div class="calendar-days-header">
                    <div class="calendar-day-header">Sun</div>
                    <div class="calendar-day-header">Mon</div>
                    <div class="calendar-day-header">Tue</div>
                    <div class="calendar-day-header">Wed</div>
                    <div class="calendar-day-header">Thu</div>
                    <div class="calendar-day-header">Fri</div>
                    <div class="calendar-day-header">Sat</div>
                </div>
                
                <asp:Repeater ID="rptCalendarDays" runat="server" OnItemDataBound="rptCalendarDays_ItemDataBound">
                    <HeaderTemplate>
                        <div class="calendar-grid">
                    </HeaderTemplate>
                    <ItemTemplate>
                        <div class="calendar-day<%# GetDayClasses(Container.DataItem) %>">
                            <div class="day-number"><%# GetDayNumber(Container.DataItem) %></div>
                            <asp:Repeater ID="rptDayEvents" runat="server">
                                <ItemTemplate>
                                    <div class="calendar-event" style="background-color: <%# GetEventColor(Container.DataItem) %>">
                                        <div class="event-title"><%# DataBinder.Eval(Container.DataItem, "Title") %></div>
                                    </div>
                                </ItemTemplate>
                            </asp:Repeater>
                        </div>
                    </ItemTemplate>
                    <FooterTemplate>
                        </div>
                    </FooterTemplate>
                </asp:Repeater>
            </div>

            <div class="calendar-sidebar">
                <h3>
                    <asp:Literal ID="ltrSelectedDate" runat="server" />
                </h3>
                <div class="event-list">
                    <asp:Repeater ID="rptEvents" runat="server" OnItemDataBound="rptEvents_ItemDataBound">
                        <ItemTemplate>
                            <div class="event-item">
                                <div class="event-color-indicator" style="background-color: <%# GetEventColor(Container.DataItem) %>"></div>
                                <div class="event-details">
                                    <h4><%# DataBinder.Eval(Container.DataItem, "Title") %></h4>
                                    <p class="event-time">
                                        <%# FormatEventTime(Container.DataItem) %>
                                    </p>
                                    <p class="event-location"><%# DataBinder.Eval(Container.DataItem, "Location") %></p>
                                    <p class="event-description"><%# DataBinder.Eval(Container.DataItem, "Description") %></p>
                                </div>
                            </div>
                        </ItemTemplate>
                        <EmptyDataTemplate>
                            <div class="no-events">No events for this day</div>
                        </EmptyDataTemplate>
                    </asp:Repeater>
                </div>
            </div>
        </div>

        <style type="text/css">
            /* Calendar Container */
            .monthly-calendar-container {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                display: flex;
                flex-wrap: wrap;
                max-width: 1200px;
                margin: 0 auto;
            }

            /* Calendar Header */
            .calendar-header {
                width: 100%;
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 20px;
                flex-wrap: wrap;
            }

            .calendar-header h2 {
                font-size: 24px;
                margin: 0;
                color: #222c36;
            }

            .month-display {
                font-size: 18px;
                font-weight: bold;
                color: #555;
            }

            /* Navigation Buttons */
            .calendar-navigation {
                display: flex;
                gap: 10px;
            }

            .nav-button {
                padding: 8px 15px;
                background-color: #222c36;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                text-decoration: none;
                display: inline-block;
            }

            .nav-button:hover {
                background-color: #eee4db;
                color: #222c36;
            }

            /* Calendar Grid */
            .calendar-body {
                width: 70%;
                margin-right: 2%;
            }

            .calendar-days-header {
                display: grid;
                grid-template-columns: repeat(7, 1fr);
                background-color: #222c36;
                color: white;
                font-weight: bold;
            }

            .calendar-day-header {
                padding: 10px;
                text-align: center;
            }

            .calendar-grid {
                display: grid;
                grid-template-columns: repeat(7, 1fr);
                grid-gap: 1px;
                background-color: #f2f2f2;
                border: 1px solid #ddd;
            }

            .calendar-day {
                min-height: 100px;
                background-color: white;
                padding: 5px;
                position: relative;
            }

            .calendar-day.other-month {
                background-color: #f9f9f9;
                color: #aaa;
            }

            .calendar-day.today {
                background-color: #eee4db;
            }

            .calendar-day.selected {
                background-color: #f0f8ff;
                border: 2px solid #007bff;
            }

            .day-number {
                position: absolute;
                top: 5px;
                right: 5px;
                font-weight: bold;
            }

            /* Calendar Events */
            .calendar-event {
                margin: 2px 0;
                padding: 3px 5px;
                border-radius: 3px;
                color: white;
                font-size: 11px;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
                margin-top: 25px;
            }

            /* Sidebar */
            .calendar-sidebar {
                width: 28%;
                background-color: #f9f9f9;
                padding: 15px;
                border-radius: 5px;
            }

            .calendar-sidebar h3 {
                margin-top: 0;
                padding-bottom: 10px;
                border-bottom: 1px solid #ddd;
                color: #222c36;
            }

            .event-list {
                max-height: 500px;
                overflow-y: auto;
            }

            .event-item {
                display: flex;
                margin-bottom: 15px;
                padding-bottom: 15px;
                border-bottom: 1px solid #eee;
            }

            .event-color-indicator {
                width: 5px;
                margin-right: 10px;
                border-radius: 2px;
            }

            .event-details {
                flex-grow: 1;
            }

            .event-details h4 {
                margin: 0 0 5px 0;
                color: #333;
            }

            .event-time, .event-location {
                margin: 0 0 5px 0;
                font-size: 13px;
                color: #666;
            }

            .event-description {
                margin: 5px 0 0 0;
                font-size: 14px;
                color: #555;
            }

            .no-events {
                color: #999;
                font-style: italic;
                padding: 10px 0;
            }

            /* Responsive Design */
            @media (max-width: 768px) {
                .calendar-body, .calendar-sidebar {
                    width: 100%;
                    margin-right: 0;
                }
                
                .calendar-sidebar {
                    margin-top: 20px;
                }
                
                .calendar-day {
                    min-height: 80px;
                }
            }
        </style>
    </ContentTemplate>
</asp:UpdatePanel>
