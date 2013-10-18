<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>TabLinker</title>
 
<link rel="stylesheet" href="css/main.css" type="text/css" />
 
<!--[if IE]>
    <script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script><![endif]-->
<!--[if lte IE 7]>
    <script src="js/IE8.js" type="text/javascript"></script><![endif]-->
<!--[if lt IE 7]>
 
    <link rel="stylesheet" type="text/css" media="all" href="css/ie6.css"/><![endif]-->
</head>
 
<body id="index" class="home">
<center>
<h1>TabLinker</h1>
%if state == 'start':
    <form action="http://lod.cedar-project.nl:8081/tablinker/upload" method="post" enctype="multipart/form-data">
    Select a file: <input type="file" name="upload" />
    <input type="submit" value="Start upload" />
    </form>
%elif state == 'uploaded':
    <form action="http://lod.cedar-project.nl:8081/tablinker/run" method="get">
    Upload OK, click to convert
    <input type="submit" value="Convert to RDF" />
    </form>
%else:
    <form action="http://lod.cedar-project.nl:8081/tablinker/download" method="get">
    TabLinkger generated {{numtriples}} triples successfully. Click to download
    <input type="submit" value="Download TTL" />
    </form>
%end
</center>
</body>
</html>
