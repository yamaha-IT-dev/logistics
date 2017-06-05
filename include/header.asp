<%
Server.ScriptTimeout=2000
'Response.Write(Server.ScriptTimeout)
'setup for Australian Date/Time
session.lcid = 2057
session.timeout = 420

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

' Constants that we need for this system
const shipmentModuleID = 1
const transferModuleID = 2
const itemMaintenanceModuleID = 3
const freightModuleID = 4
const warehouseDamageModuleID = 5
const changeoverModuleID = 6
const stockmodModuleID = 7
const warehouseReturnModuleID = 8
const graModuleID = 9
const cancelledModuleID = 11
const warehouse3thReturnModuleID = 12
const unsolicitedGoodsModuleID = 13

dim strSection

function displayDateFormatted(strDateInput)
    if IsNull(strDateInput) or strDateInput = "01/01/1900" or strDateInput = "1/1/1900"  then
        Response.Write "N/A"
    else
        Response.Write "" & WeekDayName(WeekDay(strDateInput)) & ", " & FormatDateTime(strDateInput,1) & " at " & FormatDateTime(strDateInput,3)
    end if
end function
%>
<tr>
    <td height="50" valign="top">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" height="50">
            <tr>
                <td align="left" valign="top" style="padding-left:15px; padding-top:15px;">
                    <% Response.Write WeekDayName(WeekDay(Date, vbLongDate)) & ", " & FormatDateTime(Date, vbLongDate) %>
                    <br />
                    G'day <b><%= Session("UsrUserName") %></b>! <small>(<a href="update_user.asp">Change Password</a>)</small>
                </td>
                <td align="right" style="padding-right:15px; padding-top:8px; padding-bottom:10px;">
                    <img src="images/yamaha_logo.jpg" border="0" />
                </td>
            </tr>
        </table>
    </td>
</tr>
<tr>
    <td height="26" valign="top">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="26">
            <tr>

                <!-- Shipments -->
                <%
                if session("UsrLoginRole") <> 4 and session("UsrLoginRole") <> 5 and session("UsrLoginRole") <> 7 and session("UsrLoginRole") <> 9 and session("UsrLoginRole") <> 10 and session("UsrLoginRole") <> 11 and session("UsrLoginRole") <> 12 and session("UsrLoginRole") <> 13 then
                    if strSection = "shipment" Then
                %>
                        <td class="emptyrownavigation" width="1px">&nbsp;</td>
                        <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                        <td width="95" class="selectednavigation" nowrap="nowrap"><a href="list_shipment.asp">Shipments</a></td>
                        <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td> 
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="95" class="mainnavigation" nowrap="nowrap"><a href="list_shipment.asp">Shipments</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>

                <!-- Changeover -->
                <%
                if Session("UsrLoginRole") = 1 or Session("UsrLoginRole") = 6 or Session("UsrLoginRole") = 7 or Session("UsrLoginRole") = 12 then
                    if strSection = "changeover" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="95" class="selectednavigation"><a href="list_changeover.asp">Changeover</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation"><a href="list_changeover.asp">Changeover</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>

                <!-- Stock Mods -->
                <%
                if Session("UsrLoginRole") = 1 or Session("UsrLoginRole") = 5 or Session("UsrLoginRole") = 9 or Session("UsrLoginRole") = 12 or session("UsrLoginRole") = 14 or session("UsrLoginRole") = 17 then
                    if strSection = "stock_modification" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="95" class="selectednavigation" nowrap="nowrap"><a href="list_stockmod.asp">Stock Mods</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td> 
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation" nowrap="nowrap"><a href="list_stockmod.asp">Stock Mods</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>

                <!-- GRA - Reports -->
                <%
                if Session("UsrLoginRole") = 1 or Session("UsrLoginRole") = 2 or Session("UsrLoginRole") = 6 or Session("UsrLoginRole") = 7 or Session("UsrLoginRole") = 9 or Session("UsrLoginRole") = 12 or session("UsrLoginRole") = 14 or session("UsrLoginRole") = 17 then
                    if strSection = "gra" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="95" class="selectednavigation" nowrap="nowrap"><a href="list_gra.asp">GRA</a> - <a href="list_gra_report.asp">Reports</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation" nowrap="nowrap"><a href="list_gra.asp">GRA</a> - <a href="list_gra_report.asp">Reports</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>

                <!-- FOCUS -->
                <%
                if Session("UsrLoginRole") = 1 or Session("UsrLoginRole") = 4 then
                    if strSection = "roadshow" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="100" class="selectednavigation"><a href="list_focus.asp">FOCUS</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation"><a href="list_focus.asp">FOCUS</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>

                <!-- WH Damage -->
                <%
                if Session("UsrLoginRole") = 1 or Session("UsrLoginRole") = 8 or Session("UsrLoginRole") = 14 or session("UsrLoginRole") = 17 then
                    if strSection = "damage" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="100" class="selectednavigation" nowrap="nowrap"><a href="list_damage.asp">WH Damage</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation" nowrap="nowrap"><a href="list_damage.asp">WH Damage</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>

                <!-- Cancel Order -->
                <%
                if Session("UsrLoginRole") = 1 or session("UsrLoginRole") = 14 or session("UsrLoginRole") = 17 then
                    if strSection = "cancelled" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="100" class="selectednavigation" nowrap="nowrap"><a href="list_cancelled.asp">Cancel Order</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation" nowrap="nowrap"><a href="list_cancelled.asp">Cancel Order</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>

                <!-- Freights -->
                <%
                if Session("UsrLoginRole") = 1 or Session("UsrLoginRole") = 6 then
                    if strSection = "freight" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="100" class="selectednavigation"><a href="list_freights.asp">Freights</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation"><a href="list_freights.asp">Freights</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>

                <!-- Transfers -->
                <%
                if Session("UsrLoginRole") = 1 or Session("UsrLoginRole") = 6 or Session("UsrLoginRole") = 7 or session("UsrLoginRole") = 13 or session("UsrLoginRole") = 14 or session("UsrLoginRole") = 17 then
                    if strSection = "transfer" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="100" class="selectednavigation"><a href="list_transfer.asp">Transfers</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation"><a href="list_transfer.asp">Transfers</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>

                <!-- Item Maint. -->
                <%
                if Session("UsrLoginRole") = 1 or Session("UsrLoginRole") = 2 or Session("UsrLoginRole") = 3 or Session("UsrLoginRole") = 4 or Session("UsrLoginRole") = 5 or Session("UsrLoginRole") = 6 or Session("UsrLoginRole") = 9 or Session("UsrLoginRole") = 11 or Session("UsrLoginRole") = 12 then
                    if strSection = "item" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="100" class="selectednavigation" nowrap="nowrap"><a href="list_item-maintenance.asp">Item Maint.</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation" nowrap="nowrap"><a href="list_item-maintenance.asp">Item Maint.</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>

                <!-- WH Return -->
                <%
                if Session("UsrLoginRole") = 1 or session("UsrLoginRole") = 12 or session("UsrLoginRole") = 14 or session("UsrLoginRole") = 17 then
                    if strSection = "quarantine" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="95" class="selectednavigation" nowrap="nowrap"><a href="list_warehouse-return.asp">WH Return</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation" nowrap="nowrap"><a href="list_warehouse-return.asp">WH Return</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>

                <!-- 3TH -->
                <%
                if Session("UsrLoginRole") = 1 or session("UsrLoginRole") = 12 or session("UsrLoginRole") = 14 or session("UsrLoginRole") = 17 then
                    if strSection = "3TH" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="95" class="selectednavigation" nowrap="nowrap"><a href="list_3TH.asp">3TH</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation" nowrap="nowrap"><a href="list_3TH.asp">3TH</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>

                <!-- Pack -->
                <%
                if strSection = "pack" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="95" class="selectednavigation" nowrap="nowrap"><a href="list_pack.asp">Pack</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation" nowrap="nowrap"><a href="list_pack.asp">Pack</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                end if
                %>

                <!-- Incomplete -->
                <%
                if strSection = "unsolicited" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="95" class="selectednavigation" nowrap="nowrap"><a href="list_unsolicited.asp">Incomplete</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation" nowrap="nowrap"><a href="list_unsolicited.asp">Incomplete</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                end if
                %>

                <!-- Stock Availability -->
                <%
                if Session("UsrLoginRole") = 1 or Session("UsrLoginRole") = 7 then
                    if strSection = "stockavailability" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="120" class="selectednavigation" nowrap="nowrap">Stock Availability</td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="120" class="mainnavigation" nowrap="nowrap"><a href="list_stock.asp">Stock Availability</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>

                <!-- Dashboard -->
                <%
                if Session("UsrLoginRole") = 1 or Session("UsrLoginRole") = 2 or Session("UsrLoginRole") = 9 or Session("UsrLoginRole") = 17 then
                    if strSection = "dashboard" Then
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_selected_left.gif" width="2" /></td>
                    <td width="95" class="selectednavigation" nowrap="nowrap"><a href="dashboard.asp">Dashboard</a></td>
                    <td width="2"><img src="images/menu_separator_selected_right.gif" width="2" /></td>
                <%
                    else
                %>
                    <td class="emptyrownavigation" width="1px">&nbsp;</td>
                    <td width="2"><img src="images/menu_separator_left.gif" width="2" /></td>
                    <td width="100" class="mainnavigation" nowrap="nowrap"><a href="dashboard.asp">Dashboard</a></td>
                    <td width="2"><img src="images/menu_separator_right.gif" width="2" /></td>
                <%
                    end if
                end if
                %>
                <td class="emptyrownavigation" nowrap="nowrap">&nbsp;<img src="images/forward_arrow.gif" border="0" /> <a href="default.asp?logout=y">Logout</a></td>
            </tr>
        </table>
    </td>
</tr>