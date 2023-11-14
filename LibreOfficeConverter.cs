using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace libre_office_pdf_converter_console
{
    public static class LibreOfficeConverter
    {
        public static void ConvertFileToPdf(string fileName)
        {
            Console.WriteLine();
            Console.WriteLine($"Start converting {fileName}");

            string[] wordFileTypes = new string[3] { ".docx", ".doc", ".rtf" };
            string[] excelFileTypes = new string[3] { ".xlsx", ".xls", ".xlsm" };
            string[] powerPointFileTypes = new string[2] { ".pptx", ".ppt" };

            try
            {
                string baseDir = $"{Directory.GetCurrentDirectory()}\\Files";
                string path = $"{baseDir}\\{fileName}";
                string ext = Path.GetExtension(path);   // e.g. '.docx'

                if (wordFileTypes.Contains(ext) || excelFileTypes.Contains(ext) || powerPointFileTypes.Contains(ext))
                {
                    Console.WriteLine("Starting process ...");

                    Process process = new Process();
                    ProcessStartInfo processStartInfo = new ProcessStartInfo();
                    processStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    processStartInfo.FileName = "cmd.exe";

                    // https://stackoverflow.com/questions/30349542/command-libreoffice-headless-convert-to-pdf-test-docx-outdir-pdf-is-not
                    string converter = wordFileTypes.Contains(ext)
                        ? "pdf:writer_pdf_Export"
                        : excelFileTypes.Contains(ext)
                        ? "pdf:calc_pdf_Export"
                        : powerPointFileTypes.Contains(ext)
                        ? "pdf:impress_pdf_Export"
                        : String.Empty;
                    string userProfilePath = "-env:UserInstallation=file:///C:/tmp/LibreOffice_Conversion";
                    string libreOfficeExe = @"C:\\Program Files\\LibreOffice\\program\\soffice.exe";

                    // command must start with /C so that it executes
                    // all the arguments must be enclosed in "" - otherwise the cmd.exe will split them by intervals in the path name, e.g C:\Program Files\
                    // base dir must not end with \
                    string command = $@"/C """"{libreOfficeExe}"" --headless ""{userProfilePath}"" --convert-to {converter} --outdir ""{baseDir}"" ""{path}""""";

                    Console.WriteLine(command);

                    processStartInfo.Arguments = command;
                    process.StartInfo = processStartInfo;
                    process.Start();
                    process.WaitForExit();

                    Console.WriteLine($"Saved {fileName} as pdf");
                }
                else
                {
                    Console.WriteLine("Unsupported file type!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
