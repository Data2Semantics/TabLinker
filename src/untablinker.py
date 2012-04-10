"""
Created on 19 Sep 2011
Modified on 31 Jan 2012

Authors:    Rinke Hoekstra, Laurens Rietveld
Copyright:  VU University Amsterdam, 2011/2012
License:    LGPLv3

"""
import xlwt
import glob
from rdflib import ConjunctiveGraph, Namespace, plugin
import rdflib
import re
from ConfigParser import SafeConfigParser
import logging
import os
import sys
import shutil
try:
    import rdfextras#@UnusedImport
except ImportError, e:
    print("RDF extras package is required for this script to work. Install by executing:")
    print("sudo easy_install rdfextras")
    quit()




class UnTabLinker(object):
    defaultNamespacePrefix = 'http://www.data2semantics.org/data/'
    namespaces = {
      'dcterms':Namespace('http://purl.org/dc/terms/'), 
      'skos':Namespace('http://www.w3.org/2004/02/skos/core#'), 
      'd2s':Namespace('http://www.data2semantics.org/core/'), 
      'qb':Namespace('http://purl.org/linked-data/cube#'), 
      'owl':Namespace('http://www.w3.org/2002/07/owl#')
    }

    def __init__(self, directory, config, level = logging.DEBUG):
        """TabLinker constructor
        
        Keyword arguments:
        directory -- String containing the name turtle file
        config -- Configuration object, loaded from .ini file
        level -- A logging level as defined in the logging module
        """
        self.config = config
        self.log = logging.getLogger("TabLinker")
        self.log.setLevel(level)
        
        self.graph = ConjunctiveGraph()
        self.log.debug('Loading and parsing file')
        self.graph.parse(filename, format=config.get('general', 'format'))
        
        plugin.register('sparql', rdflib.query.Processor,'rdfextras.sparql.processor', 'Processor')
        plugin.register('sparql', rdflib.query.Result,'rdfextras.sparql.query', 'SPARQLQueryResult')
        
        


    def saveFiles(self, directory):
        """
        Convert data in rdf to excel
        
        Keyword arguments:
        directory -- Directory to save files in
        """
        
        #In the future, have a sparql query here to retrieve all files in this graph
        #Not possible now, because tablinker doesnt explicitly store filename yet
        self.saveFile(directory, "converted")
        
        
    
    def saveFile(self, directory, filename):
        """
        Retrieve information from graph, and save file
        """
        self.wbk = xlwt.Workbook()
        
        
        #Retrieve sheets for this file
        
        #In the future, have a sparql query here to retrieve all sheets for this file
        #Not possible now, because tablinker doesnt explicitly store sheet name yet
        self.addSheetToXls("sheetname")
        
        self.wbk.save(directory + filename + '.xls')
    
    def addSheetToXls(self, sheetName):
        """
        Get values for this sheet, and store in excel object
        """
        self.sheet = self.wbk.add_sheet(sheetName)
        
        #Get the row IDs from the RDF set
        queryResult = self.graph.query(
            """SELECT DISTINCT ?cell ?value
               WHERE {
                  ?node <http://www.data2semantics.org/core/cell> ?cell .
                  ?node <http://www.data2semantics.org/core/value> ?value .
               }""",
            #Can't use prefix d2s. This produces parsing error (event though namespace is defined).
            #A bug in the query parser I guess
            #also, dont use [] in this query processor...
            initNs=self.namespaces
        )
        if (len(queryResult) == 0):
            self.log.error("No rows found for sheet {0}".format(sheetName))
        
        #Loop through cells and add values to excel
        for resultRow in queryResult.result:
            cell, value = resultRow
            col, row = self.cellname2index(cell)
            self.sheet.write(row, col, value)
            
        
    def cellname2index(self, cellname): 
        matches = re.search('([A-Z]*)([0-9]*)',cellname)
        if (len(matches.groups()) != 2 ):
            logging.error("Failed to parse cell name {0}. Exiting...".format(cellname))
            quit()
        col = reduce(lambda s,a:s*26+ord(a)-ord('A'), matches.group(1), 0)
        row = int(matches.group(2)) - 1
        return col,row 

def checkArg() :
    """
    Checks validity of argument (filename)
    """
    if (len(sys.argv) != 2):
        logging.error("Please provider turtlefile as parameter. exiting...")
        quit()
    if (os.path.isfile(sys.argv[1]) == False):
        logging.error("File {0} does not exist. exiting...".format(sys.argv[1]))
        quit()
    fileExtension = os.path.splitext(sys.argv[1])[1]
    if (fileExtension != '.ttl'):
        logging.error("Only ttl files are accepted. Current extension: {0}. exiting...".format(fileExtension))
        quit()

if __name__ == '__main__':
    """
    Start the un-TabLinker for ttl file
    """
    logging.basicConfig(level=logging.INFO)
    checkArg()
    filename = sys.argv[1]
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
        logging.error("Could not find configuration file. Exiting")
        quit()
        
    logging.basicConfig(level=logLevel)
    
    # Get list of annotated XLS files
    files = glob.glob(srcMask)
    logging.info("Found {0} files to convert.".format(len(files)))
    
    unLinker = UnTabLinker(filename, config, logLevel)
    
    basename = os.path.basename(filename)
    basename = re.search('(.*)\.ttl',basename).group(1)
    directory = config.get('paths', 'targetFolder') + basename + "/"
    
    if (os.path.isdir(directory)) :
        logging.debug('Output dir {0} already exists. Deleting'.format(directory))
        shutil.rmtree(directory)
    logging.debug('Creating dir {0}'.format(directory))
    os.makedirs(directory)
    unLinker.saveFiles(directory)
    
    
    logging.info("Done")
    

        
