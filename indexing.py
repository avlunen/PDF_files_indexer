#-----------------------------------------------------------------------
# Name: Indexer
# Purpose: To extract a list of words from MS Word documents in order to
# help creating an index, v2.5
#
# Author: Alexander von Lunen
#
# Created: 04/07/2019
# Copyright: (c) Alexander von Lunen 2015--2019
# Licence: Public Domain
#-----------------------------------------------------------------------
import glob
import os
import time
import io
import codecs
from topia.termextract import tag
from topia.termextract import extract
from docx import Document


# Retrieve the text of a Word document from a docx object
def getdocumenttext(document):
   return '\n'.join([paragraph.text
      for paragraph in document.paragraphs])

# main routine
def main():
   try:
      # list of index terms
      index_list = list()

      # init tagging
      tagger = tag.Tagger()
      tagger.initialize()
      extractor = extract.TermExtractor(tagger)
      #extractor.filter = extract.permissiveFilter
      #extractor.filter = extract.DefaultFilter(singleStrengthMinOccur=2)

      # get file path; you may need to customize this
      p = os.path.join('*.docx')

      # go through files
      for infile in glob.glob(p):
         # open document
         doc = Document(os.getcwd()+os.sep+infile)
         print os.getcwd()+os.sep+infile

         # get text from Word document
         text = getdocumenttext(doc)

         # tagging
         l = extractor(text)
         for item in l:
            if item[0] not in index_list:
               index_list.append(item[0])

         # close Word document
         del doc

         file = codecs.open(os.getcwd()+os.sep+'all_concordances.tsv', 'w', 'utf8')
         for row in sorted(index_list):
            file.write(row+'\t\n')
         file.close()
   finally:
      print "Done!"

if __name__ == '__main__':
   main()

