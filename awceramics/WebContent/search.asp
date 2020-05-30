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
   <header >
 <div class="container">
		 <div class="col-sm-2 col-md-2 "> 
		 </div>
	 
		<div class="col-sm-8 col-md-8 "> 
		   	<div><img src="./images/allison wiffen logo.jpg"  data-pin-nopin = "true" alt="Allison Wiffen Logo" class="img-fluid" width="75%"></a>
		   	</div>
		</div>
	     	
	    <div class="col-sm-2 col-md-2 "> 
		</div>
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

strHeader = "<h1 id=""header"">Search Results for " & Server.HTMLEncode(strSearch) & " - " & i
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
	<h2>Cufflink results</h2>
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
<a class="style-3" href="<%= rsProductSearch("itemhref") %>"><%= rsProductSearch("itemname") %></a>
<p><p>	<a class="style-3" href="<%= rsProductSearch("rangehref") %>">Range - <%= rsProductSearch("rangename") %></a>
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


	iCount = 0
%>
	</div>
	<h2>Pocket square results</h2>
<%



if i > 0 then
	rsProductSearch.MoveFirst
end if

Do While not rsProductSearch.EOF 

if rsProductSearch("type") = "square" then

if (iCount =0) then
%>
	
	<div class="row">
	<% 
end if
%>

    <div class="col-sm-4" class="col-md-4">
<div class="ranges_thumbnail"> <a href="<%= rsProductSearch("itemhref") %>"><img src="<%= rsProductSearch("imagename") %>" alt="<%= rsProductSearch("itemname") %>" ></a>
<a class="style-3" href="<%= rsProductSearch("itemhref") %>"><%= rsProductSearch("itemname") %></a>
<p><p>	<a class="style-3" href="<%= rsProductSearch("rangehref") %>">Range - <%= rsProductSearch("rangename") %></a>
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


	iCount = 0
%>
	</div>
	<h2>Brooch results</h2>
<%



if i > 0 then
	rsProductSearch.MoveFirst
end if

Do While not rsProductSearch.EOF 

if rsProductSearch("type") = "brooch" then

if (iCount =0) then
%>
	
	<div class="row">
	<% 
end if
%>

    <div class="col-sm-4" class="col-md-4">
<div class="ranges_thumbnail"> <a href="<%= rsProductSearch("itemhref") %>"><img src="<%= rsProductSearch("imagename") %>" alt="<%= rsProductSearch("itemname") %>" ></a>
<a class="style-3" href="<%= rsProductSearch("itemhref") %>"><%= rsProductSearch("itemname") %></a>
<p><p>	<a class="style-3" href="<%= rsProductSearch("rangehref") %>">Range - <%= rsProductSearch("rangename") %></a>
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


	iCount = 0
%>
	</div>
	<h2>Earring results</h2>
<%



if i > 0 then
	rsProductSearch.MoveFirst
end if

Do While not rsProductSearch.EOF 

if rsProductSearch("type") = "earrings" then

if (iCount =0) then
%>
	<div class="row">
	<% 
end if
%>

    <div class="col-sm-4" class="col-md-4">
<div class="ranges_thumbnail"> <a href="<%= rsProductSearch("itemhref") %>"><img src="<%= rsProductSearch("imagename") %>" alt="<%= rsProductSearch("itemname") %>" ></a>
<a class="style-3" href="<%= rsProductSearch("itemhref") %>"><%= rsProductSearch("itemname") %></a>
<p><p>	<a class="style-3" href="<%= rsProductSearch("rangehref") %>">Range - <%= rsProductSearch("rangename") %></a>
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


	iCount = 0
%>
	</div>
	<h2>Pendant results</h2>
<%



if i > 0 then
	rsProductSearch.MoveFirst
end if

Do While not rsProductSearch.EOF 

if rsProductSearch("type") = "pendant" then

if (iCount =0) then
%>
	<div class="row">
	<% 
end if
%>

    <div class="col-sm-4" class="col-md-4">
<div class="ranges_thumbnail"> <a href="<%= rsProductSearch("itemhref") %>"><img src="<%= rsProductSearch("imagename") %>" alt="<%= rsProductSearch("itemname") %>" ></a>
<a class="style-3" href="<%= rsProductSearch("itemhref") %>"><%= rsProductSearch("itemname") %></a>
<p><p>	<a class="style-3" href="<%= rsProductSearch("rangehref") %>">Range - <%= rsProductSearch("rangename") %></a>
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

if rsProductSearch("type") = "pendant2" then

if (iCount =0) then
%>
	
	<div class="row">
	<% 
end if
%>

    <div class="col-sm-4" class="col-md-4">
<div class="ranges_thumbnail"> <a href="<%= rsProductSearch("itemhref") %>"><img src="<%= rsProductSearch("imagename") %>" alt="<%= rsProductSearch("itemname") %>" ></a>
<a class="style-3" href="<%= rsProductSearch("itemhref") %>"><%= rsProductSearch("itemname") %></a>
<p><p>	<a class="style-3" href="<%= rsProductSearch("rangehref") %>">Range - <%= rsProductSearch("rangename") %></a>
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

if (iCount <2) then
	iCount = 0
%>
	</div>
<%
end if




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