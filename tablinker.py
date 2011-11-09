"""
Created on 19 Sep 2011

@author: hoekstra
"""
from xlutils.margins import number_of_good_cols, number_of_good_rows
from xlutils.copy import copy
from xlutils.styles import Styles
from xlrd import open_workbook, XL_CELL_TEXT, XL_CELL_EMPTY, XL_CELL_BLANK, cellnameabs, cellname
from xlwt import easyxf
import glob
from rdflib import ConjunctiveGraph, Namespace, Literal, RDF, RDFS, URIRef, XSD, BNode
import re
from ConfigParser import SafeConfigParser

#set default encoding to latin-1, to avoid encode/decode errors for special chars
#(laurens: actually don't know why encoding/decoding is not sufficient)
import sys
reload(sys)
sys.setdefaultencoding("latin-1")

config = SafeConfigParser()
config.read('config.ini')
DCTERMS = Namespace('http://purl.org/dc/terms/')
SKOS = Namespace('http://www.w3.org/2004/02/skos/core#')
D2S = Namespace('http://www.data2semantics.org/core/')
QB = Namespace('http://purl.org/linked-data/cube#')
OWL = Namespace('http://www.w3.org/2002/07/owl#')


def initGraph(scope):
    """Initialize graph

    Keyword arguments:
    string -- Scope to init graph for 
    
    Returns:
    ConjunctiveGraph
    Namespace -- Namespace for given scope
    """
    CENSUS = Namespace('http://www.data2semantics.org/data/'+scope+'/')
	
    graph = ConjunctiveGraph()

    # Bind namespaces to graph
    graph.namespace_manager.bind('dcterms',DCTERMS)
    graph.namespace_manager.bind('skos',SKOS)
    graph.namespace_manager.bind('census',CENSUS)
    graph.namespace_manager.bind('d2s',D2S)
    graph.namespace_manager.bind('qb',QB)
    graph.namespace_manager.bind('owl',OWL)
    
    graph.add((D2S['populationSize'], RDF.type, QB['MeasureProperty']))
    graph.add((D2S['populationSize'], RDFS.label, Literal('Population Size','en')))
    graph.add((D2S['populationSize'], RDFS.label, Literal('Populatie grootte','nl')))
    graph.add((D2S['populationSize'], RDFS.range, XSD.decimal))
    
    graph.add((D2S['dimension'], RDF.type, QB['DimensionProperty']))
    
    graph.add((D2S['label'], RDF.type, RDF['Property']))

    return graph, CENSUS
    
def getType(style):
    """Get type for a given excel style. Style must be prefixed by D2S

    Keyword arguments:
    string -- Style to check type for
    
    Returns:
    String -- The type of this field. In case none is found, 'unknown'
    """
    typematch = re.search('D2S\s(.*)',style)
    if typematch :
        type = typematch.group(1)
    else :
        type = 'Unknown'
    return type

def isEmpty(i,j):
    if (r_sheet.cell(i,j).ctype == XL_CELL_EMPTY or r_sheet.cell(i,j).ctype == XL_CELL_BLANK) or r_sheet.cell(i,j).value == '' :
        return True
    else :
        return False
    
def isEmptyRow(i,colns):
    for j in range(0,colns) :
        if not isEmpty(i,j):
            return False
    return True

def isEmptyColumn(j,rowns):
    for i in range(0,rowns) :
        if not isEmpty(i,j):
            return False
    return True


def getQName(names = {}):
    """
    Keyword arguments:
    mixed -- Either dict of names or string of a name.
    
    Returns:
    ConjunctiveGraph
    Namespace -- Namespace for given scope
    """
    
    qname = re.sub('\s','_',r_sheet.name)
    
    if type(names) == dict :
        for k in names :
            qname = qname + '/' + encodeString(names[k])
        return qname
    else :
        return qname + '/' + encodeString(names)
    
def encodeString(string):
    string = unicode(string)
    string = re.sub('\s|\.|\(|\)|,|:|;|\[|\]','_',string.strip()).encode('utf-8', 'ignore')
    return string
    

def getLeftWithValue(i,j,type):
    # Get value of first cell to the left of type 'type' that's not empty.
    if j == 0:
        return None, None
    
    left = r_sheet.cell(i,j-1)
    left_name = cellname(i,j-1)
    
    if isEmpty(i,j-1) :
        return getLeftWithValue(i, j-1, type)
    elif getType(styles[left].name) == type :
        return left, left_name
    else :
        return getLeftWithValue(i, j-1, type)
        



def addValue(graph, sheet_qname, source_cell_qname, source_cell_value, label=None):
    
    if not label:
        label = source_cell_value
        
    source_cell_value_qname = getQName(source_cell_value)
    graph.add((CENSUS[source_cell_value_qname],QB['dataSet'],CENSUS[sheet_qname]))
    graph.add((CENSUS[source_cell_value_qname],RDFS.label,Literal(label,'nl')))
    graph.add((CENSUS[source_cell_qname],D2S['value'],CENSUS[source_cell_value_qname]))
    
    return graph, source_cell_value_qname

def debug(i, j, msg) :
    """
    If verbose debug level is enabled, then a debug statement is printed. 
    Row/col numbers are translated into excel cellnames
    
    Keyword arguments:
    int i -- row number
    int j -- col number
    string msg -- Msg to display
    
    Returns:
    New row hierarchy dictionary
    """
    if (config.getboolean('debug', 'verbose')) :
        print cellname(i,j) + ": " + msg

def getValidRowsCols(sheet) :
    colns = number_of_good_cols(sheet)
    rowns = number_of_good_rows(sheet)
    
    # Check whether the number of good columns and rows are correct
    while isEmptyRow(rowns-1, colns) :
        rowns = rowns - 1 
    while isEmptyColumn(colns-1, rowns) :
        colns = colns - 1
    return rowns, colns

def parse(r_sheet, w_sheet, graph, CENSUS):
    rowns, colns = getValidRowsCols(r_sheet)
    print "Parsing " + str(rowns) + " rows and " + str(colns) + " cols"
    
    sheet_qname = getQName()
    
    dimcol = {}
    dimrow = {}
    rowhierarchy = {}
    
    for i in range(0,rowns):
        rowhierarchy[i] = {}
        for j in range(0, colns):
            source_cell = r_sheet.cell(i,j)
            source_cell_name = cellname(i,j)
            style = styles[source_cell].name
            type = getType(style)
 
            debug(i,j,"{}/{}: \"{}\"".format(type, source_cell_name, source_cell.value))
            
            if (type == 'HierarchicalRowHeader') :
                rowhierarchy = parseHierarchicalRowHeader(i, j, rowhierarchy)
            
            if not isEmpty(i,j) :
                source_cell_qname = getQName(source_cell_name) 
                
                graph.add((CENSUS[source_cell_qname],RDF.type,D2S[type]))
                
                if type == 'Title' :
                    graph, source_cell_value_qname = addValue(graph, sheet_qname, source_cell_qname, source_cell.value)
                    graph.add((CENSUS[sheet_qname], D2S['title'], CENSUS[source_cell_value_qname]))
                    graph.add((CENSUS[source_cell_value_qname],RDF.type,D2S['Dimension']))

                elif type == 'Property' :
                    graph, source_cell_value_qname = addValue(graph, sheet_qname, source_cell_qname, source_cell.value)
                    graph.add((CENSUS[source_cell_qname],D2S['isDimensionProperty'],CENSUS[source_cell_value_qname]))
                    graph.add((CENSUS[source_cell_value_qname],RDF.type,QB['DimensionProperty']))
                    graph.add((CENSUS[source_cell_value_qname],RDF.type,RDF['Property']))
                        
                    if j in dimcol :
                        dimcol[j].append(source_cell_value_qname)
                    else :
                        dimcol[j] = []
                        dimcol[j].append(source_cell_value_qname)     
                                       
                elif type == 'Header' :
                    graph, source_cell_value_qname = addValue(graph, sheet_qname, source_cell_qname, source_cell.value)   
                    graph.add((CENSUS[source_cell_qname],D2S['isDimension'],CENSUS[source_cell_value_qname]))
                    graph.add((CENSUS[source_cell_value_qname],RDF.type,D2S['Dimension']))
                        
                    if j in dimcol :
                        dimcol[j].append(source_cell_value_qname)
                    else :
                        dimcol[j] = []
                        dimcol[j].append(source_cell_value_qname)
                    
                elif type == 'RowHeader' :
                    graph, source_cell_value_qname = addValue(graph, sheet_qname, source_cell_qname, source_cell.value)
                    graph.add((CENSUS[source_cell_qname],D2S['isDimension'],CENSUS[source_cell_value_qname]))
                    graph.add((CENSUS[source_cell_value_qname],RDF.type,D2S['Dimension']))

                    
                    # Get the properties to use for the row headers
                    try :
                        properties = []
                        for dim_qname in dimcol[j] :
                            properties.append(dim_qname)
                    except KeyError :
                        debug(i,j, "No row dimension for cell")
                    
                    if i in dimrow :
                        dimrow[i].append((source_cell_value_qname,properties))
                    else :
                        dimrow[i] = []
                        dimrow[i].append((source_cell_value_qname,properties))
                    
                elif type == 'HierarchicalRowHeader' :
                    # Use the rowhierarchy to create a unique qname for the cell's contents, give the source_cell's original value as extra argument
                    debug(i,j,"Row hierarchy " + str(rowhierarchy[i]))
                    
                    graph, source_cell_value_qname = addValue(graph, sheet_qname, source_cell_qname, rowhierarchy[i], label=source_cell.value)
                    
                    graph.add((CENSUS[source_cell_value_qname], RDFS.comment, Literal('Copied value, original: '+ source_cell.value, 'nl')))
                        
                    # Now that we know the source cell's value qname, add a link.
                    graph.add((CENSUS[source_cell_qname], D2S['isDimension'], CENSUS[source_cell_value_qname]))
                    
                    hierarchy_items = rowhierarchy[i].items()
                    try: 
                        parent_values = dict(hierarchy_items[:-1])
                        debug(i,j, "Parent value: " + str(parent_values))
                        parent_value_qname = getQName(parent_values)
                        graph.add((CENSUS[source_cell_value_qname], SKOS['broader'], CENSUS[parent_value_qname]))
                    except :
                        debug(i,j, "Top of hierarchy")
                 
                    # Get the properties to use for the row headers
                    try :
                        properties = []
                        for dim_qname in dimcol[j] :
                            properties.append(dim_qname)
                    except KeyError :
                        debug(i,j, "No row dimension for cell")
                    
                    if i in dimrow :
                        dimrow[i].append((source_cell_value_qname,properties))
                    else :
                        dimrow[i] = []
                        dimrow[i].append((source_cell_value_qname,properties))
                    
                elif type == 'Data' :
                    observation = BNode()
                    
                    graph.add((CENSUS[source_cell_qname],D2S['isObservation'], observation))
                    graph.add((observation,RDF.type,QB['Observation']))
                    graph.add((observation,QB['dataSet'],CENSUS[sheet_qname]))
                    graph.add((observation,D2S['populationSize'],Literal(source_cell.value)))
                    
                    try :
                        for (dim_qname, properties) in dimrow[i] :
                            for p in properties:
                                graph.add((observation,D2S[p],CENSUS[dim_qname]))
                    except KeyError :
                        debug(i,j, "No row dimension for cell")
                        
                    try :
                        for dim_qname in dimcol[j] :
                            graph.add((observation,D2S['dimension'],CENSUS[dim_qname]))
                    except KeyError :
                        debug(i,j, "No row dimension for cell")
                    
    return graph


def parseHierarchicalRowHeader(i, j, rowhierarchy) :
    """
    Build up lists for hierarchical row headers
    
    Keyword arguments:
    int i -- row number
    int j -- col number
    dict rowhierarchy -- Current build row hierarchy
    
    Returns:
    New row hierarchy dictionary
    """
    source_cell = r_sheet.cell(i,j)
    source_cell_name = cellname(i,j)
    if (isEmpty(i,j) or str(source_cell.value).lower().strip() == 'id.') :
        # If the cell is empty, and a HierarchicalRowHeader, add the value of the row header above it.
        # If the cell is exactly 'id.', add the value of the row header above it. 
        try :
            rowhierarchy[i][j] = rowhierarchy[i-1][j]
            debug(i,j,"Copied from above\nRow hierarchy " + str(rowhierarchy[i]))
        except :
            rowhierarchy[i][j] = source_cell.value
            debug(i,j, "Top row, added value\nRow hierarchy: " + str(rowhierarchy[i]))
    elif str(source_cell.value).lower().startswith('id.') :
        # If the cell starts with 'id.', add the value of the row header above it, and append the rest of the cell's value.
        suffix = source_cell.value[3:]               
        try :       
            rowhierarchy[i][j] = rowhierarchy[i-1][j]+suffix
            debug(i,j, "Copied from above+suffix\nRow hierarchy " + str(rowhierarchy[i]))
        except :
            rowhierarchy[i][j] = source_cell.value
            debug(i,j, "Top row, added value\nRow hierarchy " + str(rowhierarchy[i]))
    elif not isEmpty(i,j) :
        rowhierarchy[i][j] = source_cell.value
        debug(i,j, "Added value\nRow hierarchy " + str(rowhierarchy[i]))
    return rowhierarchy
    
if __name__ == '__main__':
    # Open census data files
    fileFound = False
    for filename in glob.glob(config.get('paths', 'srcFolder') + '*_marked.xls') :
        fileFound = True
        rb = open_workbook(filename, formatting_info=True)
        
        scope = re.search('.*/(.*?)\.xls',filename).group(1)
        graph, CENSUS = initGraph(scope)
        
        #Get styles from xls file
        styles = Styles(rb)
        
        wb = copy(rb)
        
        for n in range(rb.nsheets) :
            r_sheet = rb.sheet_by_index(n)
            w_sheet = wb.get_sheet(n)
            graph = parse(r_sheet, w_sheet, graph, CENSUS)
                

        
        print "Serializing graph to file {}.ttl".format(scope)
        graph.serialize(config.get('paths', 'trgtFolder') + scope+'.ttl', format='turtle')
    
    if fileFound :
        print "Done"
    else :
        print "No files found. Path with location of marked xls files ok?"
        print "Path searched in: " + config.get('paths', 'srcFolder')
