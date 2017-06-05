<%
'setup for Australian Date/Time
session.lcid = 2057
session.timeout = 420

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
<div class="navbar navbar-default navbar-fixed-top" role="navigation">
  <div class="container" style="width:100%">
    <div class="navbar-header">
      <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-collapse"> <span class="sr-only">Toggle navigation</span> <span class="icon-bar"></span> <span class="icon-bar"></span> <span class="icon-bar"></span> </button>
      <a class="navbar-brand" href="dashboard.asp" title="Home"><img src="images/home.png"></a></div>
    <div class="navbar-collapse collapse">
      <ul class="nav navbar-nav">
        <li class="dropdown"><a href="#" class="dropdown-toggle" data-toggle="dropdown" title="Shipment">Shipment <b class="caret"></b></a>
          <ul class="dropdown-menu">
            <li><a href="add_shipment.asp">New Shipment</a></li>
            <li class="divider"></li>
            <li class="dropdown-header">LIST SHIPMENT</li>
            <li><a href="list_shipment.asp">Open</a></li>
            <li><a href="list_past-shipment.asp">Past</a></li>
          </ul>
        </li>
        <li class="dropdown"><a href="list_cancelled.asp" title="Cancelled Orders">Cancelled</a></li>
        <li class="dropdown"><a href="#" class="dropdown-toggle" data-toggle="dropdown" title="Changeover">Changeovers <b class="caret"></b></a>
          <ul class="dropdown-menu">
            <li><a href="add_changeover.asp">New Changeover</a></li>
            <li><a href="list_changeover.asp">List Changeovers</a></li>
          </ul>
        </li>
        <li class="dropdown"><a href="list_freights.asp" title="Freights">Freights</a></li>
        <li class="dropdown"><a href="#" class="dropdown-toggle" data-toggle="dropdown" title="Goods Return Authorisations">GRA <b class="caret"></b></a>
          <ul class="dropdown-menu">
            <li><a href="list_gra.asp">List Goods Return BASE</a></li>
            <li class="divider"></li>
            <li class="dropdown-header">REPORTS</li>
            <li><a href="list_gra_report.asp">Summary</a></li>
            <li><a href="list_gra_report_writeoffs.asp">Write Off</a></li>
            <li><a href="list_gra_report_exported.asp">Exported</a></li>
            <li class="divider"></li>
            <li class="dropdown-header">PALLETS</li>
            <li><a href="list_pallet.asp">List Pallets</a></li>
          </ul>
        </li>
        <li class="dropdown"><a href="#" class="dropdown-toggle" data-toggle="dropdown" title="Item Maintenance">Item Maintenance <b class="caret"></b></a>
          <ul class="dropdown-menu">
            <li><a href="add_item-maintenance.asp">New Item Maintenance</a></li>
            <li><a href="list_item-maintenance.asp">List Item Maintenance</a></li>
          </ul>
        </li>
        <li class="dropdown"><a href="#" class="dropdown-toggle" data-toggle="dropdown" title="Unsolicited Goods">Unsolicited <b class="caret"></b></a>
          <ul class="dropdown-menu">
            <li><a href="add_unsolicited.asp">New Unsolicited Goods</a></li>
            <li><a href="list_unsolicited.asp">List Unsolicited Goods</a></li>
          </ul>
        </li>
        <li class="dropdown"><a href="#" class="dropdown-toggle" data-toggle="dropdown" title="Stock Modifications">Stock Mods <b class="caret"></b></a>
          <ul class="dropdown-menu">
            <li><a href="add_stockmod.asp">New Stock Mod</a></li>
            <li><a href="list_stockmod.asp">List Stock Mod</a></li>
          </ul>
        </li>
        <li class="dropdown"><a href="#" class="dropdown-toggle" data-toggle="dropdown" title="Warehouse">Warehouse <b class="caret"></b></a>
          <ul class="dropdown-menu">
            <li class="dropdown-header">DAMAGES</li>
            <li><a href="add_damage.asp">New Warehouse Damage</a></li>
            <li><a href="list_damage.asp">List Warehouse Damages</a></li>
            <li class="divider"></li>
            <li class="dropdown-header">RETURNS</li>
            <li><a href="add_warehouse-return.asp">New Warehouse Return</a></li>
            <li><a href="list_quarantine.asp">List Warehouse Returns</a></li>
            <li class="divider"></li>
            <li class="dropdown-header">TRANSFERS</li>
            <li><a href="add_transfer.asp">New Warehouse Transfer</a></li>
            <li><a href="list_transfer.asp">List Warehouse Transfers</a></li>
            <li class="divider"></li>
            <li class="dropdown-header">PACKS</li>
            <li><a href="list_pack.asp">List Packs</a></li>
            <li class="divider"></li>
            <li class="dropdown-header">INVENTORY</li>
            <li><a href="list_inventory.asp">List Inventory</a></li>
          </ul>
        </li>
        <li class="dropdown"><a href="#" class="dropdown-toggle" data-toggle="dropdown" title="3TH">3TH <b class="caret"></b></a>
          <ul class="dropdown-menu">
            <li><a href="add_3TH.asp">New 3TH Return</a></li>
            <li><a href="list_3TH.asp">List 3TH Returns</a></li>
          </ul>
        </li>
        <li><a href="./?logout=y" title="Logout" style="color:orange">Log out</a></li>
      </ul>
    </div>
  </div>
</div>
