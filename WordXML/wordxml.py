######################################
##  ------------------------------- ##
##           WordXML Class          ##
##  ------------------------------- ##
##     Written by: David Shaeffer   ##
##  ------------------------------- ##
##      Last Edited on: 9-30-2019   ##
##  ------------------------------- ##
######################################
"""This is the WordXML Class."""

from lxml import etree as ET
from tkinter import filedialog
import os

class WordXML:
    """This is the class doc string."""
    def __init__(self, xmlfile):
     """This takes the xml file path passed to the class and parses the xml using the lxml module.""" 
     self.xmlfile = ET.parse(xmlfile).getroot()
    
    def savefile(self):
     """This function saves the xml to a new file in a user defined location."""
     try:
      """Create a new xml element tree and write the updated xml to the new tree instance."""
      tree = ET.ElementTree(self.xmlfile)
      """
      Open a file dialog box using tkinter and set the default extension to xml. 
      """
      f = filedialog.asksaveasfilename(defaultextension=".xml")
      """Write the new xml tree instance to the xml file"""
      tree.write(f, xml_declaration=True, encoding="UTF-8",
      method="xml", standalone="yes")
      """Use the OS to open the xml file"""
      os.startfile(f)
     except:
      """If the user presses cancels or it cannot save for some reason then pass."""
      print ("Could not save template.")

    def editcdp(self, fieldname, fieldtext):
     """This functions modifies the custom document properties."""
     try:
      """Parse through the xml document and find the custom document property by tag."""
      for CDP in self.xmlfile.iter('{CustomDocumentProperties}' + fieldname):
         """Insert the text (field name) passed to the function to the custom document property specified (field name)."""
         CDP.text = fieldtext
         CDP.set('updated', 'yes')
     except:
         print ("Could not modify custome document property.")
         raise
    
    def removesection(self, x):
     """This function assigns sections to all paragraphs before a page break and deletes the section 
     specified by the x function parameter - TODO better way to do this?"""
     try:
      i=0
      for body in self.xmlfile.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body'):
       """Iterate through all the paragraphs in the document body."""
       for child in body:
        """Create a new temporary section property on each paragraph and sets it equal to i."""
        child.set('section', str(i))
        """Iterate through the document to determine the number of section breaks (scetPr)."""
        for sections in child.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr'):
         """For each section break increment i."""
         i += 1
     except:
      raise 
     try:
      for body in self.xmlfile.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body'):
       """Iterate through all the paragraphs in the document body."""
       for child in body:
        """If the section value equals the section passed to the function (x) then remove all paragraphs in that section."""
        if int(child.get('section')) == x:
         child.getparent().remove(child)
     except:
      raise
    
    def removecdp(self, fieldname):
     """This functions removes the custom document property based on the properties name."""
     try:
      """Parse through the xml document and find all the data bindings."""
      for CDP in self.xmlfile.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}dataBinding'):
         """Search the xpath to determine if the databindings is the targetted custome document property."""
         if fieldname in CDP.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}xpath'):
            """Delete the custom document property up to the sdt element"""
            CDP.getparent().getparent().getparent().remove(CDP.getparent().getparent())
     except:
         print ("Could not modify custom document property.")
         raise
       
