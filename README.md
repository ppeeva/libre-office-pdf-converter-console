# libre-office-pdf-converter-console

Console application which executes a command to use LibreOffice without GUI for converting office files to pdf (doc, docx, rtf, xls, xlsx, xlsxm, ppt, pptx).

Requires LibreOffice to be installed on the machine it is running. LibreOffice is Free and Open Source Software.

The command line looks like:

"path to installed soffice.exe" --headless "-env:UserInstallation=file:///path to default user folder" --convert-to pdf:writer_pdf_Export --outdir "path to output dir" "path to the file to be converted"

The application has a folder with sample files. The converted pdf files are also saved to the same folder.
