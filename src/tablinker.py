#!/usr/bin/python

"""
Created on 19 Sep 2011
Modified on 22 Feb 2014

Authors:    Rinke Hoekstra, Laurens Rietveld, Albert Meronyo-Penyuela
Copyright:  VU University Amsterdam, 2011, 2012, 2013, 2014
License:    LGPLv3

"""
from xlutils.margins import number_of_good_cols, number_of_good_rows
from xlutils.copy import copy
from xlutils.styles import Styles
from xlrd import open_workbook, XL_CELL_EMPTY, XL_CELL_BLANK, cellname, colname
import glob
from rdflib import ConjunctiveGraph, Namespace, Literal, RDF, RDFS, BNode, URIRef
import re
from ConfigParser import SafeConfigParser
import urllib
from urlparse import urlparse
import logging
import os
import time
import datetime
#set default encoding to latin-1, to avoid encode/decode errors for special chars
#(laurens: actually don't know why encoding/decoding is not sufficient)
#(rinke: this is a specific requirment for the xlrd and xlutils packages)

import sys
reload(sys)
import traceback
sys.setdefaultencoding("utf8") #@UndefinedVariable


class TabLinker(object):
    defaultNamespacePrefix = 'http://lod.cedar-project.nl/resource/'
    annotationsNamespacePrefix = 'http://lod.cedar-project.nl/annotations/'
    namespaces = {
      'dcterms':Namespace('http://purl.org/dc/terms/'), 
      'skos':Namespace('http://www.w3.org/2004/02/skos/core#'), 
      'd2s':Namespace('http://lod.cedar-project.nl/core/'), 
      'qb':Namespace('http://purl.org/linked-data/cube#'), 
      'owl':Namespace('http://www.w3.org/2002/07/owl#')
    }
    annotationNamespaces = {
      'np':Namespace('http://www.nanopub.org/nschema#'),
      'oa':Namespace('http://www.w3.org/ns/openannotation/core/'),
      'xsd':Namespace('http://www.w3.org/2001/XMLSchema#'),
      'dct':Namespace('http://purl.org/dc/terms/')
    }

    def __init__(self, filename, config, level = logging.DEBUG):
        """TabLinker constructor
        
        Keyword arguments:
        filename -- String containing the name of the current Excel file being examined
        config -- Configuration object, loaded from .ini file
        level -- A logging level as defined in the logging module
        """
        self.config = config
        self.filename = filename
         
        self.log = logging.getLogger("TabLinker")
        self.log.setLevel(level)
        
        self.log.debug('Initializing Graphs')
        self.initGraphs()
        
        self.log.debug('Setting Scope')
        basename = os.path.basename(filename)
        basename = re.search('(.*)\.xls',basename).group(1)
        self.setScope(basename)
        
        self.log.debug('Loading Excel file {0}.'.format(filename))
        self.rb = open_workbook(filename, formatting_info=True)
        
        self.log.debug('Reading styles')
        self.styles = Styles(self.rb)
        
        self.log.debug('Copied Workbook to writable copy')
        self.wb = copy(self.rb)
        
        
    def initGraphs(self):
        """Initialize the graphs, set default namespaces, and add schema information"""
    
        self.graph = ConjunctiveGraph()
        # Create a separate graph for annotations
        self.annotationGraph = ConjunctiveGraph()
        
        self.log.debug('Adding namespaces to graphs')
        # Bind namespaces to graphs
        for namespace in self.namespaces:
            self.graph.namespace_manager.bind(namespace, self.namespaces[namespace])

        # Same for annotation graph
        for namespace in self.annotationNamespaces:
            self.annotationGraph.namespace_manager.bind(namespace, self.annotationNamespaces[namespace])
        
        self.log.debug('Adding some schema information (dimension and measure properties) ')
        self.addDataCellProperty()
                    
        self.graph.add((self.namespaces['d2s']['dimension'], RDF.type, self.namespaces['qb']['DimensionProperty']))
        
        self.graph.add((self.namespaces['d2s']['label'], RDF.type, RDF['Property']))
    
    def addDataCellProperty(self):
        """Add definition of data cell resource to graph"""

        if len(self.config.get('dataCell', 'propertyName')) > 0 :
            self.dataCellPropertyName = self.config.get('dataCell', 'propertyName')
        else :
            self.dataCellPropertyName = 'hasValue'
        
        self.graph.add((self.namespaces['d2s'][self.dataCellPropertyName], RDF.type, self.namespaces['qb']['MeasureProperty']))
        
        #Take labels from config
        if len(self.config.get('dataCell', 'labels')) > 0 :
            labels = self.config.get('dataCell', 'labels').split(':::')
            for label in labels :
                labelProperties = label.split('-->')
                if len(labelProperties[0]) > 0 and len(labelProperties[1]) > 0 :
                    self.graph.add((self.namespaces['d2s'][self.dataCellPropertyName], RDFS.label, Literal(labelProperties[1],labelProperties[0])))
                    
        if len(self.config.get('dataCell', 'literalType')) > 0 :
            self.graph.add((self.namespaces['d2s'][self.dataCellPropertyName], RDFS.range, URIRef(self.config.get('dataCell', 'literalType'))))
            
    def setScope(self, fileBasename):
        """Set the default namespace and base for all URIs of the current workbook"""
        self.fileBasename = fileBasename
        scopeNamespace = self.defaultNamespacePrefix + fileBasename + '/'
        
        # Annotations go to a different namespace
        annotationScopeNamespace = self.annotationsNamespacePrefix + fileBasename + '/'
        
        self.log.debug('Adding namespace for {0}: {1}'.format(fileBasename, scopeNamespace))
        
        self.namespaces['scope'] = Namespace(scopeNamespace)
        self.annotationNamespaces['scope'] = Namespace(annotationScopeNamespace)
        self.graph.namespace_manager.bind('', self.namespaces['scope'])
        self.annotationGraph.namespace_manager.bind('', self.annotationNamespaces['scope'])
        
    def doLink(self):
        """Start tablinker for all sheets in workbook"""
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
    
    def insideMergeBox(self, i, j):
        """
        Check if the specified cell is inside a merge box

        Arguments:
        i -- row
        j -- column

        Returns:
        True/False -- depending on whether the cell is inside a merge box
        """
        self.merged_cells = self.r_sheet.merged_cells
        for crange in self.merged_cells:
            rlo, rhi, clo, chi = crange
            if i <=  rhi - 1 and i >= rlo and j <= chi - 1 and j >= clo:
                return True
        return False
        

    def getMergeBoxCoord(self, i, j):
        """
        Get the top-left corner cell of the merge box containing the specified cell

        Arguments:
        i -- row
        j -- column

        Returns:
        (k, l) -- Coordinates of the top-left corner of the merge box
        """
        if not self.insideMergeBox(i,j):
            return (-1, -1)

        self.merged_cells = self.r_sheet.merged_cells
        for crange in self.merged_cells:
            rlo, rhi, clo, chi = crange
            if i <=  rhi - 1 and i >= rlo and j <= chi - 1 and j >= clo:
                return (rlo, clo)            
         
    def getType(self, style):
        """Get type for a given excel style. Style name must be prefixed by 'TL '
    
        Arguments:
        style -- Style (string) to check type for
        
        Returns:
        String -- The type of this field. In case none is found, 'unknown'
        """
        typematch = re.search('TL\s(.*)',style)
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
                qname = qname + '/' + self.processString(names[k])
        else :
            qname = self.sheet_qname + '/' + self.processString(names)
        
        self.log.debug('Minted new QName: {}'.format(qname))
        return qname
    

    def processString(self, string):
        """
        Remove illegal characters (comma, brackets, etc) from string, and replace it with underscore. Useful for URIs
        
        Arguments:
        string -- The string representing the value of the source cell
        
        Returns:
        processedString -- The processed string
        """
        
        return urllib.quote(re.sub('\s|\(|\)|,|\.','_',unicode(string).strip()).encode('utf-8', 'ignore'))

            
    def addValue(self, source_cell_value, altLabel=None):
        """
        Add a "value" + optional label to the graph for a cell in the source Excel sheet. The value is typically the value stored in the source cell itself, but may also be a copy of another cell (e.g. in the case of 'idem.').
        
        Arguments:
        source_cell_value -- The string representing the value of the source cell
        
        Returns:
        source_cell_value_qname -- a valid QName for the value of the source cell
        """
        source_cell_value_qname = self.getQName(source_cell_value)
        self.graph.add((self.namespaces['scope'][source_cell_value_qname],self.namespaces['qb']['dataSet'],self.namespaces['scope'][self.sheet_qname]))
        
        self.graph.add((self.namespaces['scope'][self.source_cell_qname],self.namespaces['d2s']['value'],self.namespaces['scope'][source_cell_value_qname]))
        
        # If the source_cell_value is actually a dictionary (e.g. in the case of HierarchicalRowHeaders), then use the last element of the row hierarchy as prefLabel
        # Otherwise just use the source_cell_value as prefLabel
        if type(source_cell_value) == dict :
            self.graph.add((self.namespaces['scope'][source_cell_value_qname],self.namespaces['skos'].prefLabel,Literal(source_cell_value.values()[-1],'nl')))
            
            if altLabel and altLabel != source_cell_value.values()[-1]:
                # If altLabel has a value (typically for HierarchicalRowHeaders) different from the last element in the row hierarchy, we add it as alternative label. 
                self.graph.add((self.namespaces['scope'][source_cell_value_qname],self.namespaces['skos'].altLabel,Literal(altLabel,'nl')))
        else :
            self.graph.add((self.namespaces['scope'][source_cell_value_qname],self.namespaces['skos'].prefLabel,Literal(source_cell_value,'nl')))
            
            if altLabel and altLabel != source_cell_value:
                # If altLabel has a value (typically for HierarchicalRowHeaders) different from the source_cell_value, we add it as alternative label. 
                self.graph.add((self.namespaces['scope'][source_cell_value_qname],self.namespaces['skos'].altLabel,Literal(altLabel,'nl')))
        
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

        # Get dictionary of annotations
        self.annotations = self.r_sheet.cell_note_map
        
        for i in range(0,self.rowns):
            self.rowhierarchy[i] = {}
            
            for j in range(0, self.colns):
                # Parse cell data
                self.source_cell = self.r_sheet.cell(i,j)
                self.source_cell_name = cellname(i,j)
                self.style = self.styles[self.source_cell].name
                self.cellType = self.getType(self.style)
                self.source_cell_qname = self.getQName(self.source_cell_name)
                
                self.log.debug("({},{}) {}/{}: \"{}\"". format(i,j,self.cellType, self.source_cell_name, self.source_cell.value))

                # Try to parse ints to avoid ugly _0 URIs
                try:
                    if int(self.source_cell.value) == self.source_cell.value:
                        self.source_cell.value = int(self.source_cell.value)
                except ValueError:
                    self.log.debug("(%s.%s) No parseable int" % (i,j))

                # Parse annotation (if any)
                if self.config.get('annotations', 'enabled') == "1":
                    if (i,j) in self.annotations:
                        self.parseAnnotation(i, j)

                # Parse even if empty
                if (self.cellType == 'HRowHeader') :
                    self.updateRowHierarchy(i, j)
                if self.cellType == 'Data':
                    self.parseData(i, j)
                if self.cellType == 'ColHeader' :
                    self.parseColHeader(i, j)
                if self.cellType == 'RowProperty' :
                    self.parseRowProperty(i, j)
                
                if not self.isEmpty(i,j) :
                    self.graph.add((self.namespaces['scope'][self.source_cell_qname],RDF.type,self.namespaces['d2s'][self.cellType]))
                    self.graph.add((self.namespaces['scope'][self.source_cell_qname],self.namespaces['d2s']['cell'],Literal(self.source_cell_name)))
                    #self.graph.add((self.namespaces['scope'][self.source_cell_qname],self.namespaces['d2s']['col'],Literal(colname(j))))
                    #self.graph.add((self.namespaces['scope'][self.source_cell_qname],self.namespaces['d2s']['row'],Literal(i+1)))
                    #self.graph.add((self.namespaces['scope'][self.source_cell_qname] isrow row
                    if self.cellType == 'Title' :
                        self.parseTitle(i, j)
    
                    elif self.cellType == 'RowHeader' :
                        self.parseRowHeader(i, j)
                    
                    elif self.cellType == 'HRowHeader' :
                        self.parseHierarchicalRowHeader(i, j)
                         
                    elif self.cellType == 'RowLabel' :
                        self.parseRowLabel(i, j)
        
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
        self.graph.add((self.namespaces['scope'][self.source_cell_qname], self.namespaces['d2s']['isDimension'], self.namespaces['scope'][self.source_cell_value_qname]))
        self.graph.add((self.namespaces['scope'][self.source_cell_qname], RDF.type, self.namespaces['skos'].Concept))
        
        hierarchy_items = self.rowhierarchy[i].items()
        try: 
            parent_values = dict(hierarchy_items[:-1])
            self.log.debug(i,j, "Parent value: " + str(parent_values))
            parent_value_qname = self.getQName(parent_values)
            self.graph.add((self.namespaces['scope'][self.source_cell_value_qname], self.namespaces['skos']['broader'], self.namespaces['scope'][parent_value_qname]))
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

    def parseRowLabel(self, i, j):
        """
        Create relevant triples for the cell marked as Label (i, j are row and column)
        """  
        
        self.log.debug("Parsing Row Label")
        
        # Get the QName of the HierarchicalRowHeader cell that this label belongs to, based on the rowhierarchy for this row (i)
        hierarchicalRowHeader_value_qname = self.getQName(self.rowhierarchy[i])
        
        prefLabels = self.graph.objects(self.namespaces['scope'][hierarchicalRowHeader_value_qname], self.namespaces['skos'].prefLabel)
        for label in prefLabels :
            # If the hierarchicalRowHeader QName already has a preferred label, turn it into a skos:altLabel
            self.graph.remove((self.namespaces['scope'][hierarchicalRowHeader_value_qname],self.namespaces['skos'].prefLabel,label))
            self.graph.add((self.namespaces['scope'][hierarchicalRowHeader_value_qname],self.namespaces['skos'].altLabel,label))
            self.log.debug("Turned skos:prefLabel {} for {} into a skos:altLabel".format(label, hierarchicalRowHeader_value_qname))
        
        # Add the value of the label cell as skos:prefLabel to the header cell
        self.graph.add((self.namespaces['scope'][hierarchicalRowHeader_value_qname], self.namespaces['skos'].prefLabel, Literal(self.source_cell.value, 'nl')))
            
        # Record that this source_cell_qname is the label for the HierarchicalRowHeader cell
        self.graph.add((self.namespaces['scope'][self.source_cell_qname], self.namespaces['d2s']['isLabel'], self.namespaces['scope'][hierarchicalRowHeader_value_qname]))
    
    def parseRowHeader(self, i, j) :
        """
        Create relevant triples for the cell marked as RowHeader (i, j are row and column)
        """
        rowHeaderValue = ""

        # Don't attach the cell value to the namespace if it's already a URI
        isURI = urlparse(str(self.source_cell.value))
        if isURI.scheme and isURI.netloc:
            rowHeaderValue = URIRef(self.source_cell.value)
        else:
            self.source_cell_value_qname = self.addValue(self.source_cell.value)
            rowHeaderValue = self.namespaces['scope'][self.source_cell_value_qname]

        self.graph.add((self.namespaces['scope'][self.source_cell_qname],
                        self.namespaces['d2s']['isDimension'], 
                        rowHeaderValue))
        self.graph.add((rowHeaderValue,
                        RDF.type,
                        self.namespaces['d2s']['Dimension']))
        self.graph.add((rowHeaderValue, 
                        RDF.type, 
                        self.namespaces['skos'].Concept))
        
        # Get the properties to use for the row headers
        try :
            properties = []
            for dim_qname in self.property_dimensions[j] :
                properties.append(dim_qname)
        except KeyError :
            self.log.debug("({}.{}) No properties for cell".format(i,j))
        self.row_dimensions.setdefault(i,[]).append((rowHeaderValue, properties))
        
        # Use the column dimensions dictionary to find the objects of the d2s:dimension property
        try :
            for dim_qname in self.column_dimensions[j] :
                self.graph.add((rowHeaderValue,
                                self.namespaces['d2s']['dimension'],
                                self.namespaces['scope'][dim_qname]))
        except KeyError :
            self.log.debug("({}.{}) No column dimension for cell".format(i,j))
        
        return
    
    def parseColHeader(self, i, j) :
        """
        Create relevant triples for the cell marked as Header (i, j are row and column)
        """
        if self.isEmpty(i,j):
            if self.insideMergeBox(i,j):
                k, l = self.getMergeBoxCoord(i,j)
                self.source_cell_value_qname = self.addValue(self.r_sheet.cell(k,l).value)
            else:
                return
        else:            
            self.source_cell_value_qname = self.addValue(self.source_cell.value)   

        self.graph.add((self.namespaces['scope'][self.source_cell_qname],
                        self.namespaces['d2s']['isDimension'],
                        self.namespaces['scope'][self.source_cell_value_qname]))
        self.graph.add((self.namespaces['scope'][self.source_cell_value_qname],
                        RDF.type,
                        self.namespaces['d2s']['Dimension']))
        self.graph.add((self.namespaces['scope'][self.source_cell_qname], 
                        RDF.type, 
                        self.namespaces['skos'].Concept))
        
        # Add the value qname to the column_dimensions list for that column
        self.column_dimensions.setdefault(j,[]).append(self.source_cell_value_qname)

        return
    
    def parseRowProperty(self, i, j) :
        """
        Create relevant triples for the cell marked as Property (i, j are row and column)
        """
        if self.isEmpty(i,j):
            if self.insideMergeBox(i,j):
                k, l = self.getMergeBoxCoord(i,j)
                self.source_cell_value_qname = self.addValue(self.r_sheet.cell(k,l).value)
            else:
                return
        else:
            self.source_cell_value_qname = self.addValue(self.source_cell.value)   
        self.graph.add((self.namespaces['scope'][self.source_cell_qname],self.namespaces['d2s']['isDimensionProperty'],self.namespaces['scope'][self.source_cell_value_qname]))
        self.graph.add((self.namespaces['scope'][self.source_cell_value_qname],RDF.type,self.namespaces['qb']['DimensionProperty']))
        self.graph.add((self.namespaces['scope'][self.source_cell_value_qname],RDF.type,RDF['Property']))
        
        self.property_dimensions.setdefault(j,[]).append(self.source_cell_value_qname)
        
        return
    
    def parseTitle(self, i, j) :
        """
        Create relevant triples for the cell marked as Title (i, j are row and column)
        """

        self.source_cell_value_qname = self.addValue(self.source_cell.value)
        self.graph.add((self.namespaces['scope'][self.sheet_qname], self.namespaces['d2s']['title'], self.namespaces['scope'][self.source_cell_value_qname]))
        self.graph.add((self.namespaces['scope'][self.source_cell_value_qname],RDF.type,self.namespaces['d2s']['Dimension']))
        
        return
        
        
    def parseData(self, i,j) :
        """
        Create relevant triples for the cell marked as Data (i, j are row and column)
        """
        
        if self.isEmpty(i,j) and self.config.get('dataCell', 'implicitZeros') == '0':
            return

        observation = BNode()
        
        self.graph.add((self.namespaces['scope'][self.source_cell_qname],
                        self.namespaces['d2s']['isObservation'], 
                        observation))
        self.graph.add((observation,
                        RDF.type,
                        self.namespaces['qb']['Observation']))
        self.graph.add((observation,
                        self.namespaces['qb']['dataSet'],
                        self.namespaces['scope'][self.sheet_qname]))
        if self.isEmpty(i,j) and self.config.get('dataCell', 'implicitZeros') == '1':
            self.graph.add((observation,
                            self.namespaces['d2s'][self.dataCellPropertyName],
                            Literal(0)))
        else:
            self.graph.add((observation,
                            self.namespaces['d2s'][self.dataCellPropertyName],
                            Literal(self.source_cell.value)))
        
        # Use the row dimensions dictionary to find the properties that link data values to row headers
        try :
            for (dim_qname, properties) in self.row_dimensions[i] :
                for p in properties:
                    print dim_qname
                    self.graph.add((observation,
                                    self.namespaces['d2s'][p],
                                    dim_qname))
        except KeyError :
            self.log.debug("({}.{}) No row dimension for cell".format(i,j))
        
        # Use the column dimensions dictionary to find the objects of the d2s:dimension property
        try :
            for dim_qname in self.column_dimensions[j] :
                self.graph.add((observation,
                                self.namespaces['d2s']['dimension'],
                                self.namespaces['scope'][dim_qname]))
        except KeyError :
            self.log.debug("({}.{}) No column dimension for cell".format(i,j))

    def parseAnnotation(self, i, j) :
        """
        Create relevant triples for the annotation attached to cell (i, j)
        """

        if self.config.get('annotations', 'model') == 'oa':
            # Create triples according to Open Annotation model

            body = BNode()

            self.annotationGraph.add((self.annotationNamespaces['scope'][self.source_cell_qname], 
                                      RDF.type, 
                                      self.annotationNamespaces['oa']['Annotation']
                                      ))
            self.annotationGraph.add((self.annotationNamespaces['scope'][self.source_cell_qname], 
                                      self.annotationNamespaces['oa']['hasBody'], 
                                      body
                                      ))
            self.annotationGraph.add((body,
                                      RDF.value, 
                                      Literal(self.annotations[(i,j)].text.replace("\n", " ").replace("\r", " ").replace("\r\n", " ").encode('utf-8'))
                                      ))
            self.annotationGraph.add((self.annotationNamespaces['scope'][self.source_cell_qname], 
                                      self.annotationNamespaces['oa']['hasTarget'], 
                                      self.namespaces['scope'][self.source_cell_qname]
                                      ))
            self.annotationGraph.add((self.annotationNamespaces['scope'][self.source_cell_qname], 
                                      self.annotationNamespaces['oa']['annotator'], 
                                      Literal(self.annotations[(i,j)].author.encode('utf-8'))
                                      ))
            self.annotationGraph.add((self.annotationNamespaces['scope'][self.source_cell_qname], 
                                      self.annotationNamespaces['oa']['annotated'], 
                                      Literal(datetime.datetime.fromtimestamp(os.path.getmtime(self.filename)).strftime("%Y-%m-%d"),datatype=self.annotationNamespaces['xsd']['date'])
                                      ))
            self.annotationGraph.add((self.annotationNamespaces['scope'][self.source_cell_qname], 
                                      self.annotationNamespaces['oa']['generator'], 
                                      URIRef("https://github.com/Data2Semantics/TabLinker")
                                      ))
            self.annotationGraph.add((self.annotationNamespaces['scope'][self.source_cell_qname], 
                                      self.annotationNamespaces['oa']['generated'], 
                                      Literal(datetime.datetime.now().strftime("%Y-%m-%d"), datatype=self.annotationNamespaces['xsd']['date'])
                                      ))
            self.annotationGraph.add((self.annotationNamespaces['scope'][self.source_cell_qname], 
                                      self.annotationNamespaces['oa']['modelVersion'], 
                                      URIRef("http://www.openannotation.org/spec/core/20120509.html")
                                      ))
        else:
            # Create triples according to Nanopublications model
            print "Nanopublications not implemented yet!"
            


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
    

        
