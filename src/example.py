'''
Simple example file showing how a spreadsheet can be translated to python and executed
'''
from __future__ import division
import os
import sys
from pycel.excelutil import *
from pycel.excellib import *
from pycel.excelcompiler import ExcelCompiler
from os.path import normpath,abspath

if __name__ == '__main__':
    
    dir = os.path.dirname(__file__)
    fname = os.path.join(dir, "../example/western_rater_13.6.3.xlsm")
    
    print "Loading %s..." % fname
    
    # load  & compile the file to a graph, starting from D1
    c = ExcelCompiler(filename=fname)
    
    print "Compiling..., starting from D1"
    sp = c.gen_graph('P62',sheet='Quote')

    # test evaluation
    print "Premium is %s" % sp.evaluate('Quote!P62')
    
    print "Setting Q7 to 20000"
    sp.set_value('Input!Q7',20000)
    
    print "Premium is now %s (the same should happen in Excel)" % sp.evaluate('Quote!P62')
    
    # show the graph usisng matplotlib
    #print "Plotting using matplotlib..."
    #sp.plot_graph()

    ## export the graph, can be loaded by a viewer like gephi
    #print "Exporting to gexf..."
    #sp.export_to_gexf(fname + ".gexf")
    #
    #print "Serializing to disk..."
    #sp.save_to_file(fname + ".pickle")

    #print "Done"
