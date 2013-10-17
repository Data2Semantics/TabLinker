from bottle import route, run, template, request, static_file
from tablinker import TabLinker
import logging
from ConfigParser import SafeConfigParser
import glob
import sys
import traceback
import os

@route('/tablinker/version')
def version():
    return "TabLinker version"


@route('/tablinker')
@route('/tablinker/')
def tablinker():
    return '''
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
<form action="http://lod.cedar-project.nl:8081/tablinker/upload" method="post" enctype="multipart/form-data">
  Select a file: <input type="file" name="upload" />
  <input type="submit" value="Start upload" />
</form>
</center>
</body>
</html>
    '''

@route('/tablinker/upload', method='POST')
def upload():
    # category = request.forms.get('category')
    upload = request.files.get('upload')
    name, ext = os.path.splitext(upload.filename)
    if ext not in ('.xls'):
        return 'File extension ' + ext  + ' not allowed.'

    save_path = '../input/in.xls'
    upload.save(save_path, overwrite = True) # appends upload.filename automatically
    return '''
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
<form action="http://lod.cedar-project.nl:8081/tablinker/run" method="get">
  Upload OK, click to convert
  <input type="submit" value="Convert to RDF" />
</form>
</center>
</body>
</html>
'''

@route('/tablinker/run')
def tablinker():
    logging.basicConfig(level=logging.INFO)
    logging.info('Reading configuration file')
    
    config = SafeConfigParser()
    try :
        config.read('../config.ini')
        srcMask = config.get('paths', 'srcMask')
        targetFolder = config.get('paths','targetFolder')
        verbose = config.get('debug','verbose')
        if verbose == "1" :
            logLevel = logging.DEBUG
        else :
            logLevel = logging.INFO
    except :
        logging.error("Could not find configuration file, using default settings!")
        srcMask = '../input/*_marked.xls'
        targetFolder = config.get('paths', 'targetFolder')
        logLevel = logging.DEBUG
        
    logging.basicConfig(level=logLevel)
    
    # Get list of annotated XLS files
    files = glob.glob(srcMask)
    logging.info("Found {0} files to convert.".format(len(files)))
    
    if len(files) == 0 :
        logging.error("No files found. Are you sure the path to the annotated XLS files is correct?")
        logging.info("Path searched: " + srcMask)
        quit()
    
    for filename in files :
        logging.info('Starting TabLinker for {0}'.format(filename))
        
        tLinker = TabLinker(filename, config, logLevel)
        
        logging.debug('Calling linker')
        tLinker.doLink()
        logging.debug('Done linking')

        turtleFile = targetFolder + tLinker.fileBasename +'.ttl'
        turtleFileAnnotations = targetFolder + tLinker.fileBasename +'_annotations.ttl'
        logging.info("Generated {} triples.".format(len(tLinker.graph)))
        logging.info("Serializing graph to file {}".format(turtleFile))
        try :
            fileWrite = open(turtleFile, "w")
            #Avoid rdflib writing the graph itself, as this is buggy in windows.
            #Instead, retrieve string and then write (probably more memory intensive...)
            turtle = tLinker.graph.serialize(destination=None, format=config.get('general', 'format'))
            fileWrite.writelines(turtle)
            fileWrite.close()
            
            #Annotations
            if tLinker.config.get('annotations', 'enabled') == "1":
                logging.info("Generated {} triples.".format(len(tLinker.annotationGraph)))
                logging.info("Serializing annotations to file {}".format(turtleFileAnnotations))
                fileWriteAnnotations = open(turtleFileAnnotations, "w")
                turtleAnnotations = tLinker.annotationGraph.serialize(None, format=config.get('general', 'format'))
                fileWriteAnnotations.writelines(turtleAnnotations)
                fileWriteAnnotations.close()
        except :
            logging.error("Whoops! Something went wrong in serializing to output file")
            logging.info(sys.exc_info())
            traceback.print_exc(file=sys.stdout)
            
        logging.info("Done")
    return '''
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
<form action="http://lod.cedar-project.nl:8081/tablinker/download" method="get">
  TabLinkger generated ''' + str(len(tLinker.graph)) + ''' triples successfully. Click to download
  <input type="submit" value="Download TTL" />
</form>
</center>
</body>
</html>
'''

@route('/tablinker/download')
def download():
    return static_file('in.ttl', root = '../output/', download = 'tablinker.ttl')


run(host = 'lod.cedar-project.nl', port = 8081, debug = True)

