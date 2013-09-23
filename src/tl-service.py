from bottle import route, run, template
from tablinker import TabLinker
import logging
from ConfigParser import SafeConfigParser
import glob
import sys
import traceback

@route('/tablinker/version')
def version():
    return "TabLinker version"

@route('/tablinker/')
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
    print "Done!"


run(host = 'lod.cedar-project.nl', port = 8081, debug = True)

