To run this script, just download all the files in this folder, set the FULL file path and name of the input and output files in `makePDFs.ps1`, and then run the script!

The input file is expected to be tab-separated with no first line for header names.  New line format must be Windows-style `\r\n` aka CRLF.  Column order is TEAM NAME -> PLAYER NAME -> PLAYER EMAIL -> LUNCH OPTION.  PLAYER NAME and LUNCH OPTION may be blank, though in the future names may not be.  

Make sure the following files are all located in the same directory as the script when you download it:
* BouncyCastle.Crypto.dll
* Common.Logging.Core.dll
* Common.Logging.dll
* itext.html2pdf.dll
* itext.io.dll
* itext.kernel.dll
* itext.layout.dll
* itext.styledxmlparser.dll

The C# library is iText: https://github.com/itext/itext7-dotnet, https://github.com/itext/i7n-pdfhtml. The used versions were 7.1.12 for the main library and 3.0.1 for the PDF converter since those were the latest versions that could be built successfully by me (the latest version of the main library at the time has a linking issue with one of the DLLs that prevented its use).
