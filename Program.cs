using System;

namespace libre_office_pdf_converter_console
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "sample-doc.doc";
            LibreOfficeConverter.ConvertFileToPdf(fileName);

            fileName = "sample-docx.docx";
            LibreOfficeConverter.ConvertFileToPdf(fileName);

            fileName = "sample-ppt.ppt";
            LibreOfficeConverter.ConvertFileToPdf(fileName);

            fileName = "sample-pptx.pptx";
            LibreOfficeConverter.ConvertFileToPdf(fileName);

            fileName = "sample-rtf.rtf";
            LibreOfficeConverter.ConvertFileToPdf(fileName);

            fileName = "sample-template-docx.docx";
            LibreOfficeConverter.ConvertFileToPdf(fileName);

            fileName = "sample-txt.txt";
            LibreOfficeConverter.ConvertFileToPdf(fileName);

            fileName = "sample-xls.xls";
            LibreOfficeConverter.ConvertFileToPdf(fileName);

            fileName = "sample-xlsm.xlsm";
            LibreOfficeConverter.ConvertFileToPdf(fileName);

            fileName = "sample-xlsx.xlsx";
            LibreOfficeConverter.ConvertFileToPdf(fileName);
        }
    }
}
