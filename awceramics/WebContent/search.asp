<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="utf-8">  <meta name="keywords" content="gifts,presents,jewellery,cufflinks,clothing,weddings,men's cufflinks,formalwear,menswear,accessories, designer cufflinks, cuff links,UK,unusual cufflinks,interesting cufflinks,brooches, pendants, earrings, clocks, london,newcastle,glasgow, edinburgh,graffiti,urban,cars,classic cars,father's day, mother's day,wedding gifts"><meta name="description" content="Gorgeous hand-made ceramic jewellery made in Gargrave, North Yorkshire by Allison Wiffen."><meta http-equiv="pragma" content="no-cache"><meta name="revisit-after" content="14 days"><meta name="classification" content="shopping">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- The above 3 meta tags *must* come first in the head; any other head 
         content must come *after* these tags -->
    <title>Allison Wiffen Ceramics - Designer Jewellery</title>
	<!-- Bootstrap -->
     <link href="css/bower_components/bootstrap/dist/css/bootstrap.min.css" rel="stylesheet">
	  <link href="css/bower_components/bootstrap/dist/css/bootstrap-theme.min.css" rel="stylesheet">
    <link href="css/bower_components/font-awesome/css/font-awesome.min.css" rel="stylesheet">
   
	<link href="css/mystyles.css" rel="stylesheet">
  
    <link href="css/bootstrap-social.css" rel="stylesheet">
	


    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
</head>

<body>
	
	  <div id="includeMenu"></div>	
	  
   <header class="jumbotron jumbotron-fluid">

	 <div class="container">
    	<h1>Allison Wiffen Ceramics</h1>
    	<p class="lead">Miniature works of art.</p>
        </div>
		
    </header>

    <div class="container">

<ol class="breadcrumb">
  <li><a href="index.htm">Home</a></li>
  <li class="active">Search Results</li>
</ol>



   <div class="container">

			

<table class="search">


<%
Dim adoCon
Dim rsProductSearch
Dim rsProductCount
Dim rsAddQuery
Dim strSQL
Dim strSearch
Dim strOriginalSearch
Dim i
Dim wordArray
Dim strWhere
Dim strHeader
Dim iCount

iCount = 0 

strSearch = Request.Form("searchtext")
strOriginalSearch = strSearch
strSearch = Replace(strSearch, chr(39), "")


if strSearch <> "" then 
	wordArray = Split(strSearch)
	For each item in wordArray
		strWhere = strWhere & " lcase(keywords) like '%" & lcase(item) & "%' and"
	Next
	strWhere = Left(strWhere, len(strWhere)-4)
else 
	strWhere = "1 = 2"
end if


Set adoCon = Server.CreateObject("ADODB.Connection")
DSNtest="DRIVER={Microsoft Access Driver (*.mdb)}; "
DSNtest=dsntest & "DBQ=" & Server.MapPath("productDB.mdb")
adoCon.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("productDB.mdb")

set rsAddQuery = Server.createObject("ADODB.Recordset")
rsAddQuery.CursorType = 2
rsAddQuery.LockType = 3
strSQL = "select keywords from tbSearchQuery;"
rsAddQuery.Open strSQl, adoCon
rsAddQuery.AddNew
rsAddQuery.Fields("keywords") = strOriginalSearch
rsAddQuery.Update

Set rsProductCount = Server.CreateObject("ADODB.Recordset")
strSQL = "select count(*) from tbProductsearch where " & strWhere
rsProductCount.Open strSQL, adoCon

i = rsProductCount(0)

strHeader = "<h1 id=""header"">Search Results for " & strSearch & " - " & i
if i = 1 then
	strHeader = strHeader & " match </h1>"
else
	strHeader = strHeader & " matches </h1>"
end if 
Response.Write strHeader

Set rsProductSearch = Server.CreateObject("ADODB.Recordset")
strSQL = "select imagename, itemname, itemhref, rangename, rangehref, type from tbProductsearch where " & strWhere
rsProductSearch.Open strSQL, adoCon


%>
<p>
<% 


Do While not rsProductSearch.EOF 

if rsProductSearch("type") = "cufflink" then

if (iCount =0) then
%>

	<div class="row">
	<% 
end if
%>

    <div class="col-sm-4" class="col-md-4">
<div class="ranges_thumbnail"> <a href="<%= rsProductSearch("itemhref") %>"><img src="<%= rsProductSearch("imagename") %>" alt="<%= rsProductSearch("itemname") %>" ></a>
<a class="menu" href="<%= rsProductSearch("itemhref") %>"><%= rsProductSearch("itemname") %></a>
<p><p>	<a class="menu" href="<%= rsProductSearch("rangehref") %>">Range - <%= rsProductSearch("rangename") %></a>
</div>
</div>
<%
if (iCount =2) then
	iCount = 0
%>
	</div>
	<% 
else 
	iCount=iCount+1
end if

 
end if
rsProductSearch.MoveNext


Loop 


if i > 0 then
	rsProductSearch.MoveFirst
end if

Do While not rsProductSearch.EOF 
if rsProductSearch("type") = "square" then
%>
<tr>
<td><a href="<%= rsProductSearch("itemhref") %>"><img src="<%= rsProductSearch("imagename") %>" alt="<%= rsProductSearch("itemname") %>"  width="230" height="230"></a></td>
<td><a class="menu" href="<%= rsProductSearch("itemhref") %>"><%= rsProductSearch("itemname") %></a></td>
	<td><a class="menu" href="<%= rsProductSearch("rangehref") %>"><%= rsProductSearch("rangename") %></a></td>
</tr>
<% 
end if
rsProductSearch.MoveNext
Loop 

if i > 0 then
	rsProductSearch.MoveFirst
end if

Do While not rsProductSearch.EOF 
if rsProductSearch("type") = "brooch" then
%>
<tr>
<td><a href="<%= rsProductSearch("itemhref") %>"><img src="<%= rsProductSearch("imagename") %>" alt="<%= rsProductSearch("itemname") %>"  width="230" height="230"></a></td>
<td><a class="menu" href="<%= rsProductSearch("itemhref") %>"><%= rsProductSearch("itemname") %></a></td>
	<td><a class="menu" href="<%= rsProductSearch("rangehref") %>"><%= rsProductSearch("rangename") %></a></td>
</tr>
<% 
end if
rsProductSearch.MoveNext
Loop 

if i > 0 then
	rsProductSearch.MoveFirst
end if

Do While not rsProductSearch.EOF 
if rsProductSearch("type") = "earrings" then
%>
<tr>
<td><a href="<%= rsProductSearch("itemhref") %>"><img src="<%= rsProductSearch("imagename") %>" alt="<%= rsProductSearch("itemname") %>"  width="230" height="230"></a></td>
<td><a class="menu" href="<%= rsProductSearch("itemhref") %>"><%= rsProductSearch("itemname") %></a></td>
	<td><a class="menu" href="<%= rsProductSearch("rangehref") %>"><%= rsProductSearch("rangename") %></a></td>
</tr>
<% 
end if
rsProductSearch.MoveNext
Loop 

if i > 0 then
	rsProductSearch.MoveFirst
end if

Do While not rsProductSearch.EOF 
if rsProductSearch("type") = "pendant" then
%>
<tr>
<td><a href="<%= rsProductSearch("itemhref") %>"><img src="<%= rsProductSearch("imagename") %>" alt="<%= rsProductSearch("itemname") %>"  width="230" height="345"></a></td>
<td><a class="menu" href="<%= rsProductSearch("itemhref") %>"><%= rsProductSearch("itemname") %></a></td>
	<td><a class="menu" href="<%= rsProductSearch("rangehref") %>"><%= rsProductSearch("rangename") %></a></td>
</tr>
<% 
end if
rsProductSearch.MoveNext
Loop 

if i > 0 then
	rsProductSearch.MoveFirst
end if

Do While not rsProductSearch.EOF 
if rsProductSearch("type") = "pendant2" then
%>
<tr>
<td><a href="<%= rsProductSearch("itemhref") %>"><img src="<%= rsProductSearch("imagename") %>" alt="<%= rsProductSearch("itemname") %>"  width="350" height="350"></a></td>
<td><a class="menu" href="<%= rsProductSearch("itemhref") %>"><%= rsProductSearch("itemname") %></a></td>
	<td><a class="menu" href="<%= rsProductSearch("rangehref") %>"><%= rsProductSearch("rangename") %></a></td>
</tr>
<% 
end if
rsProductSearch.MoveNext
Loop %>
</tr>
</tbody>
</table>
<%
rsAddQuery.Close
rsProductCount.Close
rsProductSearch.Close
set rsProductCount = nothing
Set rsProductsearch = nothing
Set rsAddQuery = nothing
Set adoCon = nothing

%>
<p>
</div>
</div>

    <div id="includeFooter"></div>
         

    <!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
	
	<script src="css/bower_components/jquery/dist/jquery.min.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="css/bower_components/bootstrap/dist/js/bootstrap.min.js"></script>
	
	 <script> 
	    $(function(){
	      $("#includeFooter").load("footer.html"); 
	      $("#includeMenu").load("menu.html"); 
	    });
    </script> 	

</body>

</html>