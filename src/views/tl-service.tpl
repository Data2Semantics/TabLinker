<!DOCTYPE html>
<html>
  <head>
    <title>TabLinker</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="">
    <meta name="author" content="">
    <link rel="shortcut icon" href="/img/favicon.ico">
    <!-- Bootstrap -->
    <link href="/css/bootstrap.min.css" rel="stylesheet" media="screen">
    <!-- Bootstrap core CSS -->
    <link href="/css/bootstrap.css" rel="stylesheet">
    <!-- Custom styles for this template -->
    <link href="/css/starter-template.css" rel="stylesheet">

    <!-- HTML5 shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="../../assets/js/html5shiv.js"></script>
      <script src="../../assets/js/respond.min.js"></script>
    <![endif]-->
  </head>

  <body>

    <div class="navbar navbar-inverse navbar-fixed-top">
      <div class="container">
        <div class="navbar-header">
          <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-collapse">
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
          </button>
          <a class="navbar-brand" href="/tablinker">TabLinker as-a-service</a>
        </div>
        <div class="collapse navbar-collapse">
          <ul class="nav navbar-nav">
            <li class="active"><a href="/tablinker">Home</a></li>
            <li><a href="https://github.com/Data2Semantics/TabLinker" target="_blank">GitHub</a></li>
            <li><a href="mailto:cedar@cedar-project.nl">Contact</a></li>
          </ul>
        </div><!--/.nav-collapse -->
      </div>
    </div>

    <div class="container">

      <div class="starter-template">

	<img src="/img/tablinker-logo-150dpi.png">
	<p>foobar foo bar baz</p>

	%if state == 'start':
	<form action="/tablinker/upload" method="post" enctype="multipart/form-data">
	  Select a file: <input type="file" name="upload" />
	  <input type="submit" value="Start upload" />
	</form>

	%elif state == 'uploaded':
	<form action="/tablinker/run" method="get">
	  <p>Upload OK</p>
	  <input type="submit" value="Convert to RDF" />
	</form>

	%else:
	<form action="/tablinker/download" method="get">
	  <p>TabLinkger generated {{numtriples}} triples successfully</p>
	  <input type="submit" value="Download TTL" />
	</form>
	%end
    
      </div>

    </div>

    <!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script src="//code.jquery.com/jquery.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="/js/bootstrap.min.js"></script>

  </body>
</html>
