# PDF-Figure-Extractor
Goal: extract all figures and related captions from research paper pdfs in batch.

How to obtain the input text file from EndNote:
1. Select files you would like to include from the main page
2. Go to Main Menu: File > Export > Choose save file as type: Text Only; Choose Output Style as "J Catalysis"; Select "Export Selected References".

How to obtian the PDF files from EndNote:
1. Nothing needed (Just go by default). When you "find reference" for each entry, the pdf file is auto-attached to the master folder of Endnote.

How would the output be?
1. The output will be a giant word file where each figure + caption pair is preceded by its citation name.
2. Optionally saves a CSV mapping of PDF → citation → match score.

How does the code know which PDF files corresponding to which citation we asked? Because I have more PDF files in Endnote than I want, but I only wanted to download those I want.
