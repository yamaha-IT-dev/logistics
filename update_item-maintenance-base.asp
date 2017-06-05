<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsItemMaintenance.asp " -->
<%
    'Save the item_id to session
    session("item_id") = Request.Form("item_id")

    'Variable to hold the result of query
    dim queryResult

    'Run the itemCreatedInBase function which uses the item_id held in session
    queryResult = itemCreatedInBase

    'Return the result of the query
    Response.Write(queryResult)
%>