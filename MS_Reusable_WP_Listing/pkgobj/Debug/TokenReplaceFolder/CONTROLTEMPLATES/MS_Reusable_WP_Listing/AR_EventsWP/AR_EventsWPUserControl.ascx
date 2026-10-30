<%@ Assembly Name="MS_Reusable_WP_Listing, Version=1.0.0.0, Culture=neutral, PublicKeyToken=496cdbb56b5c3083" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="AR_EventsWPUserControl.ascx.cs" Inherits="MS_Reusable_WP_Listing.AR_EventsWP.AR_EventsWPUserControl" %>

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css" />

<asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
    <Triggers>
        <asp:AsyncPostBackTrigger ControlID="btnPrevMonth" EventName="Click" />
        <asp:AsyncPostBackTrigger ControlID="btnNextMonth" EventName="Click" />
    </Triggers>
    <ContentTemplate>
        <div class="monthly-calendar-container" style="direction: rtl;">
            <div class="calendar-title" style="direction: rtl;">
                <div class="title-bar" style="margin-right: 0; margin-left: 10px;"></div>
                <h2><%# GetCalendarTitle() %></h2>
            </div>
            <div class="calendar-sidebar">
                <div class="event-list">
                    <asp:Panel ID="pnlEvents" runat="server">
                        <asp:Repeater ID="rptEvents" runat="server" OnItemDataBound="rptEvents_ItemDataBound">
                            <ItemTemplate>
                                <div class="event-item" style="direction: rtl;">
                                    <div class="date-box" style="margin-right: 0; margin-left: 15px;">
                                        <div class="date-day">
                                            <%# FormatArabicDay(DataBinder.Eval(Container.DataItem, "StartDate")) %>
                                        </div>
                                        <div class="date-month">
                                            <%# FormatArabicMonthAbbrev(DataBinder.Eval(Container.DataItem, "StartDate")) %>
                                        </div>
                                    </div>
                                    <div class="event-details">
                                        <h4><%# DataBinder.Eval(Container.DataItem, "Title") %></h4>
                                        <p class="event-description"><%# TruncateDescription(DataBinder.Eval(Container.DataItem, "Description") != null ? DataBinder.Eval(Container.DataItem, "Description").ToString() : string.Empty) %></p>
                                        <a href="<%# GetEventDetailsUrl(Container.DataItem) %>" class="read-more">المزيد...</a>
                                    </div>
                                </div>
                            </ItemTemplate>
                        </asp:Repeater>
                        <asp:Panel ID="pnlNoEvents" runat="server" CssClass="no-events" Visible="false">
                            لا توجد فعاليات لهذا اليوم
                        </asp:Panel>
                    </asp:Panel>
                </div>
            </div>

            <div class="calendar-body">
                <div class="calendar-month-display">
                    <div class="month-number"><asp:Literal ID="litMonthNumber" runat="server"></asp:Literal></div>
                    <div class="month-name">
                        <div><asp:Literal ID="litMonthNameDisplay" runat="server"></asp:Literal></div>
                        <div><asp:Literal ID="litYearDisplay" runat="server"></asp:Literal></div>
                    </div>
                </div>
                <div class="month-navigation" style="direction: rtl;">
                    <asp:LinkButton ID="btnPrevMonth" runat="server" CssClass="month-nav-button" OnClick="btnPrevMonth_Click">
                        <i class="fas fa-chevron-right"></i>
                    </asp:LinkButton>
                    <div class="current-month"><asp:Literal ID="litCurrentMonth" runat="server"></asp:Literal></div>
                    <asp:LinkButton ID="btnNextMonth" runat="server" CssClass="month-nav-button" OnClick="btnNextMonth_Click">
                        <i class="fas fa-chevron-left"></i>
                    </asp:LinkButton>
                </div>
                <div class="calendar-days-header" style="direction: rtl;">
                    <div class="calendar-day-header">الأحد</div>
                    <div class="calendar-day-header">الاثنين</div>
                    <div class="calendar-day-header">الثلاثاء</div>
                    <div class="calendar-day-header">الأربعاء</div>
                    <div class="calendar-day-header">الخميس</div>
                    <div class="calendar-day-header">الجمعة</div>
                    <div class="calendar-day-header">السبت</div>
                </div>

                <asp:Repeater ID="rptCalendarDays" runat="server" OnItemDataBound="rptCalendarDays_ItemDataBound">
                    <HeaderTemplate>
                        <div class="calendar-grid" style="direction: rtl;">
                    </HeaderTemplate>
                    <ItemTemplate>
                        <div class="calendar-day<%# GetDayClasses(Container.DataItem) %>">
                            <asp:LinkButton ID="lnkDay" runat="server" CssClass="day-number" CommandName="SelectDay" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "Date", "{0:yyyy-MM-dd}") %>' OnCommand="DayLink_Command"><%# GetDayNumber(Container.DataItem) %></asp:LinkButton>
                            <asp:Repeater ID="rptDayEvents" runat="server">
                                <ItemTemplate>
                                    <asp:LinkButton ID="lnkEvent" runat="server" CssClass="calendar-event" style="background-color: <%# GetEventColor(Container.DataItem) %>" CommandName="SelectDay" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "StartDate", "{0:yyyy-MM-dd}") %>' OnCommand="DayLink_Command">
                                        <div class="event-title"><%# DataBinder.Eval(Container.DataItem, "Title") %></div>
                                    </asp:LinkButton>
                                </ItemTemplate>
                            </asp:Repeater>
                        </div>
                    </ItemTemplate>
                    <FooterTemplate>
                        </div>
                    </FooterTemplate>
                </asp:Repeater>
            </div>
        </div>

        <style type="text/css">
            /* Month Navigation */
            .calendar-month-display {
                background-color: #a9793f; /* Havana color */
                color: white;
                padding: 15px;
            }

            .month-number {
                font-size: 2.5rem;
                font-weight: bold;
                margin-right: 10px;
                display: inline-block;
                vertical-align: middle;
            }

            .month-name {
                font-size: 1rem;
                line-height: 1.2;
                display: inline-block;
                vertical-align: middle;
            }

            .month-navigation {
                display: flex;
                align-items: center;
                justify-content: space-between;
                background-color: #333;
                color: white;
                padding: 10px 15px;
            }

            .month-nav-button {
                color: white;
                background-color: transparent;
                border: none;
                cursor: pointer;
                font-size: 1rem;
            }

            .current-month {
                font-size: 1rem;
                font-weight: 500;
            }

            /* Calendar Title */
            .calendar-title {
                width: 100%;
                display: flex;
                align-items: center;
                margin-bottom: 15px;
                padding: 0;
            }

            .title-bar {
                width: 6px;
                height: 30px;
                background-color: #a9793f; /* Havana color */
                margin-right: 10px;
                border-radius: 2px;
            }

            .calendar-title h2 {
                margin: 0;
                font-size: 16px;
                font-weight: bold;
                color: #333;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }

            /* Calendar Container */
            .monthly-calendar-container {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                display: flex;
                flex-wrap: wrap;
                max-width: 100%;
                margin: 0 auto;
                font-size: 14px;
                border: 1px solid #ddd;
                background-color: white;
            }

            /* Navigation and Month Display */
            .calendar-navigation {
                width: 100%;
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 20px;
                flex-wrap: wrap;
            }

            .month-display {
                font-size: 18px;
                font-weight: bold;
                color: #555;
                margin-left: auto;
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
                width: 75%;
                margin-left: 2%;
            }

            .calendar-days-header {
                display: grid;
                grid-template-columns: repeat(7, 1fr);
                background-color: white;
                color: black;
                font-weight: bold;
            }

            .calendar-day-header {
                padding: 12px;
                text-align: right;
                font-size: 16px;
            }

            .calendar-grid {
                display: grid;
                grid-template-columns: repeat(7, 1fr);
                grid-auto-rows: 1fr;
                grid-gap: 1px;
                background-color: #f2f2f2;
                border: 1px solid #ddd;
                min-height: 600px;
            }

            .calendar-day {
                height: 120px;
                background-color: white;
                padding: 8px;
                position: relative;
                overflow: hidden;
                box-sizing: border-box;
            }

            .calendar-day.other-month {
                background-color: #f9f9f9;
                color: #aaa;
            }

            .calendar-day.today {
                background-color: #eee4db;
            }

            .calendar-day.selected {
                background-color: #e6f2ff;
                border: 2px solid #0066cc;
                box-shadow: 0 0 8px rgba(0, 102, 204, 0.3);
                z-index: 1;
                position: relative;
            }

            .day-number {
                position: absolute;
                top: 5px;
                right: 5px;
                font-weight: bold;
                background: none;
                border: none;
                cursor: pointer;
                color: inherit;
                text-decoration: none;
                padding: 4px 8px;
                border-radius: 50%;
                font-size: 16px;
            }

            .day-number:hover {
                background-color: #f0f0f0;
            }

            /* Calendar Events */
            .calendar-event {
                margin: 2px 0;
                padding: 4px 6px;
                border-radius: 4px;
                color: black;
                font-size: 12px;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
                margin-top: 25px;
                display: block;
                text-decoration: none;
                cursor: pointer;
                border: none;
                width: calc(100% - 4px);
                text-align: left;
                font-weight: bold;
                box-shadow: 0 1px 2px rgba(0,0,0,0.2);
                max-height: 20px;
                line-height: 1.2;
            }

            .calendar-event:hover {
                opacity: 0.85;
                box-shadow: 0 2px 5px rgba(0,0,0,0.3);
                transform: translateY(-1px);
            }

            /* Sidebar */
            .calendar-sidebar {
                width: 23%;
                background-color: #f9f9f9;
                padding: 15px;
                border-radius: 5px;
            }

            /* Date Display */
            .date-box {
                width: 60px;
                height: 60px;
                background-color: #cccccc; /* Gray color from image */
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                margin-right: 15px;
                flex-shrink: 0;
            }

            .date-day {
                font-size: 24px;
                font-weight: normal;
                color: #333;
                margin-bottom: 0;
                line-height: 1;
            }

            .date-month {
                font-size: 16px;
                font-weight: bold;
                color: #000;
                text-transform: uppercase;
                line-height: 1;
            }

            .event-list {
                max-height: 500px;
                overflow-y: auto;
            }

            .event-item {
                display: flex;
                margin-bottom: 25px;
                padding: 0;
                border-bottom: none;
                background-color: transparent;
                align-items: flex-start;
            }

            .event-details {
                flex-grow: 1;
                display: flex;
                flex-direction: column;
            }

            .event-details h4 {
                margin: 0 0 5px 0;
                color: #333;
                font-size: 16px;
                font-weight: bold;
            }

            .event-time, .event-location {
                margin: 0 0 3px 0;
                font-size: 12px;
                color: #777;
            }

            .event-description {
                margin: 5px 0 5px 0;
                font-size: 14px;
                color: #555;
                line-height: 1.4;
                max-width: 100%;
                overflow-wrap: break-word;
                word-break: normal;
            }

            .read-more {
                display: inline-block;
                margin-top: 5px;
                color: #c68c53; /* Orange/gold color from image */
                font-size: 14px;
                text-decoration: none;
            }

            .no-events {
                color: #999;
                font-style: italic;
                padding: 10px 0;
            }

            /* Responsive Design */
            @media (max-width: 768px) {
                .monthly-calendar-container {
                    flex-direction: column;
                }

                .calendar-body, .calendar-sidebar {
                    width: 100%;
                    margin-left: 0;
                    margin-right: 0;
                }

                .calendar-body {
                    margin-top: 20px;
                    order: 2;
                }

                .calendar-sidebar {
                    order: 1;
                }

                .calendar-day {
                    min-height: 80px;
                }
            }
        </style>
    </ContentTemplate>
</asp:UpdatePanel>
