"""
Created on 19 Sep 2011

@author: hoekstra
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

    def __init__(self, filename):
        self.log = logging.getLogger(__name__)
        self.log.setLevel(logging.DEBUG)
        
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
        """Initialize graph
    
        Keyword arguments:
        string -- Scope to init graph for 
        
        Returns:
        ConjunctiveGraph
        Namespace -- Namespace for given scope
        """
    
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
            self.parse()
    
    ###
    #    Utility Functions
    ### 
         
    def getType(self, style):
        """Get type for a given excel style. Style name must be prefixed by 'D2S'
    
        Keyword arguments:
        string -- Style to check type for
        
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
        """Check if cell is empty.
        
        Keyword arguments:
        i -- Column
        j -- Row
        """
        if (self.r_sheet.cell(i,j).ctype == XL_CELL_EMPTY or self.r_sheet.cell(i,j).ctype == XL_CELL_BLANK) or self.r_sheet.cell(i,j).value == '' :
            return True
        else :
            return False
        
    def isEmptyRow(self, i, colns):
        for j in range(0,colns) :
            if not self.isEmpty(i,j):
                return False
        return True
    
    def isEmptyColumn(self, j, rowns):
        for i in range(0,rowns) :
            if not self.isEmpty(i,j):
                return False
        return True
    
    def getValidRowsCols(self) :
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
    
    
    
    
    def getQName(self, names = {}):
        """
        Keyword arguments:
        mixed -- Either dict of names or string of a name.
        
        Returns:
        QName -- a valid QName for the dict or string
        """
        
        if type(names) == dict :
            for k in names :
                qname = self.sheet_qname + '/' + urllib.quote(re.sub('\s','_',unicode(names[k]).strip()).encode('utf-8', 'ignore'))
        else :
            qname = self.sheet_qname + '/' + urllib.quote(re.sub('\s','_',unicode(names).strip()).encode('utf-8', 'ignore'))
        
        self.log.debug('Minted new QName: {}'.format(qname))
        return qname
        

            
    
    
    
    def addValue(self, source_cell_value, label=None):
        if not label:
            label = source_cell_value
            
        source_cell_value_qname = self.getQName(source_cell_value)
        self.graph.add((self.SCOPE[source_cell_value_qname],self.QB['dataSet'],self.SCOPE[self.sheet_qname]))
        self.graph.add((self.SCOPE[source_cell_value_qname],RDFS.label,Literal(label,'nl')))
        self.graph.add((self.SCOPE[self.source_cell_qname],self.D2S['value'],self.SCOPE[source_cell_value_qname]))
        
        return source_cell_value_qname
    

    

    
#    def appendItemInDict(self, dictionary, key, value) :
#        if key not in dictionary : 
#            dictionary[key] = []
#        dictionary[key].append(value)
#        return dictionary
    
    def parse(self):
        
        self.log.info("Parsing {0} rows and {1} columns.".format(self.rowns,self.colns))
        
        self.dimcol = {}
        self.dimrow = {}
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
                    #Always update headerlist, and always parse hierarchical row header, even if it doesn't contain data
                    self.updateRowHierarchy(i, j)
                    self.parseHierarchicalRowHeader(i, j)
                
                if not self.isEmpty(i,j) :
                    self.graph.add((self.SCOPE[self.source_cell_qname],RDF.type,self.D2S[self.cellType]))
                    
                    if self.cellType == 'Title' :
                        self.parseTitle(i, j)
    
                    elif self.cellType == 'Property' :
                        self.parseProperty(i, j)
                                           
                    elif self.cellType == 'Header' :
                        self.parseHeader(i, j)
                       
                    elif self.cellType == 'RowHeader' :
                        self.parseRowHeader(i, j)
                        
                    elif self.cellType == 'Data' :
                        self.parseData(i, j)
        
        self.log.info("Done parsing...")

    def updateRowHierarchy(self, i, j) :
        """
        Build up lists for hierarchical row headers
        
        Keyword arguments:
        int i -- row number
        int j -- col number
        dict rowhierarchy -- Current build row hierarchy
        
        Returns:
        New row hierarchy dictionary
        """
        if (self.isEmpty(i,j) or str(self.source_cell.value).lower().strip() == 'id.') :
            # If the cell is empty, and a HierarchicalRowHeader, add the value of the row header above it.
            # If the cell is exactly 'id.', add the value of the row header above it. 
            try :
                self.rowhierarchy[i][j] = self.rowhierarchy[i-1][j]
                self.log.debug("({},{}) Copied from above\nRow hierarchy: {}".format(i,j,self.rowhierarchy[i]))
            except :
                self.rowhierarchy[i][j] = self.source_cell.value
                self.log.debug("({},{}) Top row, added value\nRow hierarchy: {}".format(i,j,self.rowhierarchy[i]))
        elif str(self.source_cell.value).lower().startswith('id.') or str(self.source_cell.value).lower().startswith('id '):
            # If the cell starts with 'id.', add the value of the row header above it, and append the rest of the cell's value.
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

        
        # Use the rowhierarchy to create a unique qname for the cell's contents, give the source_cell's original value as extra argument
        self.log.debug("Parsing HierarchicalRowHeader")
        
        self.source_cell_value_qname = self.addValue(self.rowhierarchy[i], label=self.source_cell.value)
        
        self.graph.add((self.SCOPE[self.source_cell_value_qname], RDFS.comment, Literal('Copied value, original: '+ self.source_cell.value, 'nl')))
            
        # Now that we know the source cell's value qname, add a link.
        self.graph.add((self.SCOPE[self.source_cell_qname], self.D2S['isDimension'], self.SCOPE[self.source_cell_value_qname]))
        
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
            for dim_qname in self.dimcol[j] :
                properties.append(dim_qname)
        except KeyError :
            self.log.debug(i,j, "No row dimension for cell")

        self.dimrow.setdefault(i, []).append((self.source_cell_value_qname, properties))

    
    def parseRowHeader(self, i, j) :
        self.source_cell_value_qname = self.addValue(self.source_cell.value)
        self.graph.add((self.SCOPE[self.source_cell_qname],self.D2S['isDimension'],self.SCOPE[self.source_cell_value_qname]))
        self.graph.add((self.SCOPE[self.source_cell_value_qname],RDF.type,self.D2S['Dimension']))
        # Get the properties to use for the row headers
        try :
            properties = []
            for dim_qname in self.dimcol[j] :
                properties.append(dim_qname)
        except KeyError :
            self.log.debug(i,j, "No row dimension for cell")
        self.dimrow.setdefault(i,[]).append((self.source_cell_value_qname, properties))
        
        return
    
    def parseHeader(self, i, j) :
        self.source_cell_value_qname = self.addValue(self.source_cell.value)   
        self.graph.add((self.SCOPE[self.source_cell_qname],self.D2S['isDimension'],self.SCOPE[self.source_cell_value_qname]))
        self.graph.add((self.SCOPE[self.source_cell_value_qname],RDF.type,self.D2S['Dimension']))
        
        self.dimcol.setdefault(j,[]).append(self.source_cell_value_qname)

        return
    
    def parseProperty(self, i, j) :
        self.source_cell_value_qname = self.addValue(self.source_cell.value)
        
        self.graph.add((self.SCOPE[self.source_cell_qname],self.D2S['isDimensionProperty'],self.SCOPE[self.source_cell_value_qname]))
        self.graph.add((self.SCOPE[self.source_cell_value_qname],RDF.type,self.QB['DimensionProperty']))
        self.graph.add((self.SCOPE[self.source_cell_value_qname],RDF.type,RDF['Property']))
        
        self.dimcol.setdefault(j,[]).append(self.source_cell_value_qname)
        
        return
    
    def parseTitle(self, i, j) :

        self.source_cell_value_qname = self.addValue(self.source_cell.value)
        self.graph.add((self.SCOPE[self.sheet_qname], self.D2S['title'], self.SCOPE[self.source_cell_value_qname]))
        self.graph.add((self.SCOPE[self.source_cell_value_qname],RDF.type,self.D2S['Dimension']))
        
        return
        
        
    def parseData(self, i,j) :

        observation = BNode()
        
        self.graph.add((self.SCOPE[self.source_cell_qname],self.D2S['isObservation'], observation))
        self.graph.add((observation,RDF.type,self.QB['Observation']))
        self.graph.add((observation,self.QB['dataSet'],self.SCOPE[self.sheet_qname]))
        self.graph.add((observation,self.D2S['populationSize'],Literal(self.source_cell.value)))
        
        try :
            for (dim_qname, properties) in self.dimrow[i] :
                for p in properties:
                    self.graph.add((observation,self.D2S[p],self.SCOPE[dim_qname]))
        except KeyError :
            self.log.debug(i,j, "No row dimension for cell")
            
        try :
            for dim_qname in self.dimcol[j] :
                self.graph.add((observation,self.D2S['dimension'],self.SCOPE[dim_qname]))
        except KeyError :
            self.log.debug(i,j, "No row dimension for cell")

    



if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG)
    
    logging.debug('Reading configuration file')
    
    config = SafeConfigParser()
    config.read('../config.ini')
    
    # Open census data files
    files = glob.glob(config.get('paths', 'srcFolder') + '*_marked.xls')
    
    if len(files) == 0 :
        logging.WARNING("No files found. Are you sure the path to the annotated XLS files is correct?")
        logging.info("Path searched: " + config.get('paths', 'srcFolder'))
    
    for filename in files :
        
        logging.debug('Starting TabLinker for {0}'.format(filename))
        tl = TabLinker(filename)
        
        logging.debug('Calling linker')
        tl.doLink()
        logging.debug('Done linking')

        turtleFile = config.get('paths', 'trgtFolder') + tl.scope +'.ttl'
        logging.info("Serializing graph to file {}".format(turtleFile))
        tl.graph.serialize(turtleFile, format='turtle')
        logging.info("Done")
    

        
