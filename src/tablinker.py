"""
Created on 19 Sep 2011
Modified on 31 Jan 2012

Authors:    Rinke Hoekstra, Laurens Rietveld
Copyright:  VU University Amsterdam, 2011/2012
License:    LGPLv3

"""
from xlutils.margins import number_of_good_cols, number_of_good_rows
from xlutils.copy import copy
from xlutils.styles import Styles
from xlrd import open_workbook, XL_CELL_EMPTY, XL_CELL_BLANK, cellname
import glob
from rdflib import ConjunctiveGraph, Namespace, Literal, RDF, RDFS, XSD, BNode
import re
from ConfigParser import SafeConfigParser
import urllib
import logging
#set default encoding to latin-1, to avoid encode/decode errors for special chars
#(laurens: actually don't know why encoding/decoding is not sufficient)
#(rinke: this is a specific requirment for the xlrd and xlutils packages)

import sys
reload(sys)
sys.setdefaultencoding("latin-1") #@UndefinedVariable




class TabLinker(object):


    DCTERMS = Namespace('http:/g/purl.org/dc/terms/')
    SKOS = Namespace('http://www.w3.org/2004/02/skos/core#')
    D2S = Namespace('http://www.data2semantics.org/core/')
    QB = Namespace('http://purl.org/linked-data/cube#')
    OWL = Namespace('http://www.w3.org/2002/07/owl#')

    def __init__(self, filename, level = logging.DEBUG):
        """TabLinker constructor
        
        Keyword arguments:
        filename -- String containing the name of the current Excel file being examined
        level -- A logging level as defined in the logging module
        """
        self.log = logging.getLogger("TabLinker")
        self.log.setLevel(level)
        
        self.log.debug('Initializing Graph')
        self.initGraph()
        
        self.log.debug('Setting Scope')
        scope = re.search('.*/(.*?)\.xls',filename).group(1)
        self.setScope(scope)
        
        self.log.debug('Loading Excel file {0}.'.format(filename))
        self.rb = open_workbook(filename, formatting_info=True)
        
        self.log.debug('Reading styles')
        self.styles = Styles(self.rb)
        
        self.log.debug('Copied Workbook to writable copy')
        self.wb = copy(self.rb)
        
        
    def initGraph(self):
        """Initialize the graph, set default namespaces, and add schema information"""
    
        self.graph = ConjunctiveGraph()
        
        self.log.debug('Adding namespaces to graph')
        # Bind namespaces to graph
        self.graph.namespace_manager.bind('dcterms',self.DCTERMS)
        self.graph.namespace_manager.bind('skos',self.SKOS)
        self.graph.namespace_manager.bind('d2s',self.D2S)
        self.graph.namespace_manager.bind('qb',self.QB)
        self.graph.namespace_manager.bind('owl',self.OWL)
        
        self.log.debug('Adding some schema information (dimension and measure properties) ')
        self.graph.add((self.D2S['populationSize'], RDF.type, self.QB['MeasureProperty']))
        self.graph.add((self.D2S['populationSize'], RDFS.label, Literal('Population Size','en')))
        self.graph.add((self.D2S['populationSize'], RDFS.label, Literal('Populatie grootte','nl')))
        self.graph.add((self.D2S['populationSize'], RDFS.range, XSD.decimal))
        
        self.graph.add((self.D2S['dimension'], RDF.type, self.QB['DimensionProperty']))
        
        self.graph.add((self.D2S['label'], RDF.type, RDF['Property']))
    
    def setScope(self, scope):
        """Set the default namespace and base for all URIs of the current workbook"""
        self.scope = scope
        scopens = 'http://www.data2semantics.org/data/'+scope+'/'
        
        self.log.debug('Adding namespace for {0}: {1}'.format(scope, scopens))
        
        self.SCOPE = Namespace(scopens)
        self.graph.namespace_manager.bind('',self.SCOPE)
        
    def doLink(self):
        self.log.info('Starting TabLinker for all sheets in workbook')
        
        for n in range(self.rb.nsheets) :
            self.log.debug('Starting with sheet {0}'.format(n))
            self.r_sheet = self.rb.sheet_by_index(n)
            self.w_sheet = self.wb.get_sheet(n)
            
            self.rowns, self.colns = self.getValidRowsCols()
                 
            self.sheet_qname = urllib.quote(re.sub('\s','_',self.r_sheet.name))
            self.log.debug('Base for QName generator set to: {0}'.format(self.sheet_qname))
            
            self.log.debug('Starting parser')
            self.parseSheet()
    
    ###
    #    Utility Functions
    ### 
         
    def getType(self, style):
        """Get type for a given excel style. Style name must be prefixed by 'D2S '
    
        Arguments:
        style -- Style (string) to check type for
        
        Returns:
        String -- The type of this field. In case none is found, 'unknown'
        """
        typematch = re.search('D2S\s(.*)',style)
        if typematch :
            cellType = typematch.group(1)
        else :
            cellType = 'Unknown'
        return cellType
    
    def isEmpty(self, i,j):
        """Check whether cell is empty.
        
        Arguments:
        i -- row
        j -- column
        
        Returns:
        True/False -- depending on whether the cell is empty
        """
        if (self.r_sheet.cell(i,j).ctype == XL_CELL_EMPTY or self.r_sheet.cell(i,j).ctype == XL_CELL_BLANK) or self.r_sheet.cell(i,j).value == '' :
            return True
        else :
            return False
        
    def isEmptyRow(self, i, colns):
        """
        Determine whether the row 'i' is empty by iterating over all its cells
        
        Arguments:
        i     -- The index of the row to be checked.
        colns -- The number of columns to be checked
        
        Returns:
        true  -- if the row is empty
        false -- if the row is not empty
        """
        for j in range(0,colns) :
            if not self.isEmpty(i,j):
                return False
        return True
    
    def isEmptyColumn(self, j, rowns ):
        """
        Determine whether the column 'j' is empty by iterating over all its cells
        
        Arguments:
        j     -- The index of the column to be checked.
        rowns -- The number of rows to be checked
        
        Returns:
        true  -- if the column is empty
        false -- if the column is not empty
        """
        for i in range(0,rowns) :
            if not self.isEmpty(i,j):
                return False
        return True
    
    def getValidRowsCols(self) :
        """
        Determine the number of non-empty rows and columns in the Excel sheet
        
        Returns:
        rowns -- number of rows
        colns -- number of columns
        """
        colns = number_of_good_cols(self.r_sheet)
        rowns = number_of_good_rows(self.r_sheet)
        
        # Check whether the number of good columns and rows are correct
        while self.isEmptyRow(rowns-1, colns) :
            rowns = rowns - 1 
        while self.isEmptyColumn(colns-1, rowns) :
            colns = colns - 1
            
        self.log.debug('Number of rows with content:    {0}'.format(rowns))
        self.log.debug('Number of columns with content: {0}'.format(colns))
        return rowns, colns
    
    def getQName(self, names):
        """
        Create a valid QName from a string or dictionary of names
        
        Arguments:
        names -- Either dictionary of names or string of a name.
        
        Returns:
        qname -- a valid QName for the dictionary or string
        """
        
        if type(names) == dict :
            qname = self.sheet_qname
            for k in names :
                qname = qname + '/' + urllib.quote(re.sub('\s','_',unicode(names[k]).strip()).encode('utf-8', 'ignore'))
        else :
            qname = self.sheet_qname + '/' + urllib.quote(re.sub('\s','_',unicode(names).strip()).encode('utf-8', 'ignore'))
        
        self.log.debug('Minted new QName: {}'.format(qname))
        return qname
        

            
    def addValue(self, source_cell_value, altLabel=None):
        """
        Add a "value" + optional label to the graph for a cell in the source Excel sheet. The value is typically the value stored in the source cell itself, but may also be a copy of another cell (e.g. in the case of 'idem.').
        
        Arguments:
        source_cell_value -- The string representing the value of the source cell
        
        Returns:
        source_cell_value_qname -- a valid QName for the value of the source cell
        """
        source_cell_value_qname = self.getQName(source_cell_value)
        self.graph.add((self.SCOPE[source_cell_value_qname],self.QB['dataSet'],self.SCOPE[self.sheet_qname]))
        
        self.graph.add((self.SCOPE[self.source_cell_qname],self.D2S['value'],self.SCOPE[source_cell_value_qname]))
        
        # If the source_cell_value is actually a dictionary (e.g. in the case of HierarchicalRowHeaders), then use the last element of the row hierarchy as prefLabel
        # Otherwise just use the source_cell_value as prefLabel
        if type(source_cell_value) == dict :
            self.graph.add((self.SCOPE[source_cell_value_qname],self.SKOS.prefLabel,Literal(source_cell_value.values()[-1],'nl')))
            
            if altLabel and altLabel != source_cell_value.values()[-1]:
                # If altLabel has a value (typically for HierarchicalRowHeaders) different from the last element in the row hierarchy, we add it as alternative label. 
                self.graph.add((self.SCOPE[source_cell_value_qname],self.SKOS.altLabel,Literal(altLabel,'nl')))
        else :
            self.graph.add((self.SCOPE[source_cell_value_qname],self.SKOS.prefLabel,Literal(source_cell_value,'nl')))
            
            if altLabel and altLabel != source_cell_value:
                # If altLabel has a value (typically for HierarchicalRowHeaders) different from the source_cell_value, we add it as alternative label. 
                self.graph.add((self.SCOPE[source_cell_value_qname],self.SKOS.altLabel,Literal(altLabel,'nl')))
        
        

            
                    
        
        return source_cell_value_qname
    

    


    
    def parseSheet(self):
        """
        Parses the currently selected sheet in the workbook, takes no arguments. Iterates over all cells in the Excel sheet and produces relevant RDF Triples. 
        """
        self.log.info("Parsing {0} rows and {1} columns.".format(self.rowns,self.colns))
        
        self.column_dimensions = {}
        self.property_dimensions = {}
        self.row_dimensions = {}
        self.rowhierarchy = {}
        
        for i in range(0,self.rowns):
            self.rowhierarchy[i] = {}
            
            for j in range(0, self.colns):
                self.source_cell = self.r_sheet.cell(i,j)
                self.source_cell_name = cellname(i,j)
                self.style = self.styles[self.source_cell].name
                self.cellType = self.getType(self.style)
                self.source_cell_qname = self.getQName(self.source_cell_name)
     
                self.log.debug("({},{}) {}/{}: \"{}\"". format(i,j,self.cellType, self.source_cell_name, self.source_cell.value))
                
                if (self.cellType == 'HierarchicalRowHeader') :
#                    self.graph.add((self.SCOPE[self.source_cell_qname],RDF.type,self.D2S[self.cellType])) 
                    
                    #Always update headerlist even if it doesn't contain data
                    self.updateRowHierarchy(i, j)
                   
                
                if not self.isEmpty(i,j) :
                    self.graph.add((self.SCOPE[self.source_cell_qname],RDF.type,self.D2S[self.cellType]))
                    
                    if self.cellType == 'Title' :
                        self.parseTitle(i, j)
    
                    elif self.cellType == 'Property' :
                        self.parseProperty(i, j)
                                           
                    elif self.cellType == 'ColHeader' :
                        self.parseColHeader(i, j)
                       
                    elif self.cellType == 'RowHeader' :
                        self.parseRowHeader(i, j)
                    
                    elif self.cellType == 'HierarchicalRowHeader' :
                        self.parseHierarchicalRowHeader(i, j)
                         
                    elif self.cellType == 'Label' :
                        self.parseLabel(i, j)
                        
                    elif self.cellType == 'Data' :
                        self.parseData(i, j)
        
        self.log.info("Done parsing...")

    def updateRowHierarchy(self, i, j) :
        """
        Build up lists for hierarchical row headers. Cells marked as hierarchical row header are often empty meaning that their intended value is stored somewhere else in the Excel sheet.
        
        Keyword arguments:
        int i -- row number
        int j -- col number
        
        Returns:
        New row hierarchy dictionary
        """
        if (self.isEmpty(i,j) or str(self.source_cell.value).lower().strip() == 'id.') :
            # If the cell is empty, and a HierarchicalRowHeader, add the value of the row header above it.
            # If the cell above is not in the rowhierarchy, don't do anything.
            # If the cell is exactly 'id.', add the value of the row header above it. 
            try :
                self.rowhierarchy[i][j] = self.rowhierarchy[i-1][j]
                self.log.debug("({},{}) Copied from above\nRow hierarchy: {}".format(i,j,self.rowhierarchy[i]))
            except :
                # REMOVED because of double slashes in uris
                # self.rowhierarchy[i][j] = self.source_cell.value
                self.log.debug("({},{}) Top row, added nothing\nRow hierarchy: {}".format(i,j,self.rowhierarchy[i]))
        elif str(self.source_cell.value).lower().startswith('id.') or str(self.source_cell.value).lower().startswith('id '):
            # If the cell starts with 'id.', add the value of the row  above it, and append the rest of the cell's value.
            suffix = self.source_cell.value[3:]               
            try :       
                self.rowhierarchy[i][j] = self.rowhierarchy[i-1][j]+suffix
                self.log.debug("({},{}) Copied from above+suffix\nRow hierarchy {}".format(i,j,self.rowhierarchy[i]))
            except :
                self.rowhierarchy[i][j] = self.source_cell.value
                self.log.debug("({},{}) Top row, added value\nRow hierarchy {}".format(i,j,self.rowhierarchy[i]))
        elif not self.isEmpty(i,j) :
            self.rowhierarchy[i][j] = self.source_cell.value
            self.log.debug("({},{}) Added value\nRow hierarchy {}".format(i,j,self.rowhierarchy[i]))
        return self.rowhierarchy
    
    def parseHierarchicalRowHeader(self, i, j) :
        """
        Create relevant triples for the cell marked as HierarchicalRowHeader (i, j are row and column)
        """
        
        # Use the rowhierarchy to create a unique qname for the cell's contents, give the source_cell's original value as extra argument
        self.log.debug("Parsing HierarchicalRowHeader")
        
        self.source_cell_value_qname = self.addValue(self.rowhierarchy[i], altLabel=self.source_cell.value)
        
            
        # Now that we know the source cell's value qname, add a d2s:isDimension link and the skos:Concept type
        self.graph.add((self.SCOPE[self.source_cell_qname], self.D2S['isDimension'], self.SCOPE[self.source_cell_value_qname]))
        self.graph.add((self.SCOPE[self.source_cell_qname], RDF.type, self.SKOS.Concept))
        
        hierarchy_items = self.rowhierarchy[i].items()
        try: 
            parent_values = dict(hierarchy_items[:-1])
            self.log.debug(i,j, "Parent value: " + str(parent_values))
            parent_value_qname = self.getQName(parent_values)
            self.graph.add((self.SCOPE[self.source_cell_value_qname], self.SKOS['broader'], self.SCOPE[parent_value_qname]))
        except :
            self.log.debug(i,j, "Top of hierarchy")
     
        # Get the properties to use for the row headers
        try :
            properties = []
            for dim_qname in self.property_dimensions[j] :
                properties.append(dim_qname)
        except KeyError :
            self.log.debug("({}.{}) No row dimension for cell".format(i,j))

        self.row_dimensions.setdefault(i, []).append((self.source_cell_value_qname, properties))

    def parseLabel(self, i, j):
        """
        Create relevant triples for the cell marked as Label (i, j are row and column)
        """  
        
        self.log.debug("Parsing Label")
        
        # Get the QName of the HierarchicalRowHeader cell that this label belongs to, based on the rowhierarchy for this row (i)
        hierarchicalRowHeader_value_qname = self.getQName(self.rowhierarchy[i])
        
        prefLabels = self.graph.objects(self.SCOPE[hierarchicalRowHeader_value_qname], self.SKOS.prefLabel)
        for label in prefLabels :
            # If the hierarchicalRowHeader QName already has a preferred label, turn it into a skos:altLabel
            self.graph.remove((self.SCOPE[hierarchicalRowHeader_value_qname],self.SKOS.prefLabel,label))
            self.graph.add((self.SCOPE[hierarchicalRowHeader_value_qname],self.SKOS.altLabel,label))
            self.log.info("Turned skos:prefLabel {} for {} into a skos:altLabel".format(label, hierarchicalRowHeader_value_qname))
        
        # Add the value of the label cell as skos:prefLabel to the header cell
        self.graph.add((self.SCOPE[hierarchicalRowHeader_value_qname], self.SKOS.prefLabel, Literal(self.source_cell.value, 'nl')))
            
        # Record that this source_cell_qname is the label for the HierarchicalRowHeader cell
        self.graph.add((self.SCOPE[self.source_cell_qname], self.D2S['isLabel'], self.SCOPE[hierarchicalRowHeader_value_qname]))
    
    def parseRowHeader(self, i, j) :
        """
        Create relevant triples for the cell marked as RowHeader (i, j are row and column)
        """
        self.source_cell_value_qname = self.addValue(self.source_cell.value)
        self.graph.add((self.SCOPE[self.source_cell_qname],self.D2S['isDimension'],self.SCOPE[self.source_cell_value_qname]))
        self.graph.add((self.SCOPE[self.source_cell_value_qname],RDF.type,self.D2S['Dimension']))
        self.graph.add((self.SCOPE[self.source_cell_qname], RDF.type, self.SKOS.Concept))
        
        # Get the properties to use for the row headers
        try :
            properties = []
            for dim_qname in self.property_dimensions[j] :
                properties.append(dim_qname)
        except KeyError :
            self.log.debug("({}.{}) No properties for cell".format(i,j))
        self.row_dimensions.setdefault(i,[]).append((self.source_cell_value_qname, properties))
        
        # Use the column dimensions dictionary to find the objects of the d2s:dimension property
        try :
            for dim_qname in self.column_dimensions[j] :
                self.graph.add((self.SCOPE[self.source_cell_value_qname],self.D2S['dimension'],self.SCOPE[dim_qname]))
        except KeyError :
            self.log.debug("({}.{}) No column dimension for cell".format(i,j))
        
        return
    
    def parseColHeader(self, i, j) :
        """
        Create relevant triples for the cell marked as Header (i, j are row and column)
        """
        self.source_cell_value_qname = self.addValue(self.source_cell.value)   
        self.graph.add((self.SCOPE[self.source_cell_qname],self.D2S['isDimension'],self.SCOPE[self.source_cell_value_qname]))
        self.graph.add((self.SCOPE[self.source_cell_value_qname],RDF.type,self.D2S['Dimension']))
        self.graph.add((self.SCOPE[self.source_cell_qname], RDF.type, self.SKOS.Concept))
        
        # Add the value qname to the column_dimensions list for that column
        self.column_dimensions.setdefault(j,[]).append(self.source_cell_value_qname)

        return
    
    def parseProperty(self, i, j) :
        """
        Create relevant triples for the cell marked as Property (i, j are row and column)
        """
        self.source_cell_value_qname = self.addValue(self.source_cell.value)
        
        self.graph.add((self.SCOPE[self.source_cell_qname],self.D2S['isDimensionProperty'],self.SCOPE[self.source_cell_value_qname]))
        self.graph.add((self.SCOPE[self.source_cell_value_qname],RDF.type,self.QB['DimensionProperty']))
        self.graph.add((self.SCOPE[self.source_cell_value_qname],RDF.type,RDF['Property']))
        
        self.property_dimensions.setdefault(j,[]).append(self.source_cell_value_qname)
        
        return
    
    def parseTitle(self, i, j) :
        """
        Create relevant triples for the cell marked as Title (i, j are row and column)
        """

        self.source_cell_value_qname = self.addValue(self.source_cell.value)
        self.graph.add((self.SCOPE[self.sheet_qname], self.D2S['title'], self.SCOPE[self.source_cell_value_qname]))
        self.graph.add((self.SCOPE[self.source_cell_value_qname],RDF.type,self.D2S['Dimension']))
        
        return
        
        
    def parseData(self, i,j) :
        """
        Create relevant triples for the cell marked as Data (i, j are row and column)
        """

        observation = BNode()
        
        self.graph.add((self.SCOPE[self.source_cell_qname],self.D2S['isObservation'], observation))
        self.graph.add((observation,RDF.type,self.QB['Observation']))
        self.graph.add((observation,self.QB['dataSet'],self.SCOPE[self.sheet_qname]))
        self.graph.add((observation,self.D2S['populationSize'],Literal(self.source_cell.value)))
        
        # Use the row dimensions dictionary to find the properties that link data values to row headers
        try :
            for (dim_qname, properties) in self.row_dimensions[i] :
                for p in properties:
                    self.graph.add((observation,self.D2S[p],self.SCOPE[dim_qname]))
        except KeyError :
            self.log.debug("({}.{}) No row dimension for cell".format(i,j))
        
        # Use the column dimensions dictionary to find the objects of the d2s:dimension property
        try :
            for dim_qname in self.column_dimensions[j] :
                self.graph.add((observation,self.D2S['dimension'],self.SCOPE[dim_qname]))
        except KeyError :
            self.log.debug("({}.{}) No column dimension for cell".format(i,j))

    



if __name__ == '__main__':
    """
    Start the TabLinker for every file specified in the configuration file (../config.ini)
    """
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
        targetFolder = '../output/'
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
        
        tl = TabLinker(filename, logLevel)
        
        logging.debug('Calling linker')
        tl.doLink()
        logging.debug('Done linking')

        turtleFile = targetFolder + tl.scope +'.ttl'
        logging.info("Serializing graph to file {}".format(turtleFile))
        try :
            tl.graph.serialize(turtleFile, format='turtle')
        except :
            logging.error("Whoops! Something went wrong in serializing to output file")
            logging.info(sys.exc_info())
            
        logging.info("Done")
    

        
