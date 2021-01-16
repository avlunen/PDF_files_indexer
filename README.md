# PDF files indexer
[Created July 2019]

I had created a Python script to index a group of MS Word files in an earlier project: BUBU. For a new book project, I had to create/extend that script. In the previous book project, the publisher was happy to take a MS Word file with the generated index inside; we had to deliver this as part of our MS.

This time round, with a different publisher, we had to send in an unindexed MS, they sent us back the final proofs, then the final MS; from which we had to generate the index with the page numbers from the PDF. This meant while the previous script would still work, I now had to take that index list generated with the previous script, but instead of creating a concordance file for Word, I needed to export a list of the index words as plain text. This plain text file is then used by the new script to search the PDF, store the page numbers it has found each search term (a.k.a. an entry from the index file), and output this compiled index (with page numbers) as MS Word file.

The first file indexing.py is a slightly modified version from the previous project. Run this first to create a list of indexable terms (as mentioned in my previous project, you likely will have to do some heavy editing), which is outputted as TSV file, with one column with the terms and an empty second column. Edit that file as you see fit by cleaning it up; then use the second, empty column to create merged terms. For example, you may want to group terms by their geography:

Africa	Africa
African civilisation	Africa: civilisation
African history	Africa: history

But you can leave the second column empty. Once you're happy with your index list, run indexpdf.py on it. This will scan the PDF files (change the path in the code) and generate an MS Word file with the final index.

Be sure to read the source code to understand what these scripts are doing.
