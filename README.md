# tda-xls
Compile 'xsl' files (actually in excel xml format) created by 'Excel to PDF Converter' from PDF exports from TDA application into a single xsl file usable in SPSS and other analysis software.

Since the 'Excel to PDF Converter' used does not use the actual xls format, but the excel xml format, I created a module named 'xlxmlrd'. This has the exact same interface as xlrd, which I originally intended to use, and allows it be used similarly.

You'll see a lot of hacky stuff here, but this entire project is a hack, anyway; so, just enjoy it!

Olusegun Ojo
