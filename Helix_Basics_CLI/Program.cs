using System;
using System.Configuration;
using System.IO;
using CLI_Wrapper;
using System.Windows.Forms;
using System.Reflection;
using iTextSharp.text.pdf;
using System.Diagnostics;
    

namespace Helix_Basics_CLI
{
    internal class Program
    {
        internal static void Main(string[] args)
        {
        ProgramStart:
            releaseExcel();
            LogHandeler.ModuleLogFile("---", "---", "---", "Program", "Main", "Function started", null, "15", null, null, null);
            //string[] args = new string[0];
            /* ************************************************
             * Parameters for processing
             * 1. Input Folder
             * 2. Output Folder
             * 3. Processed Folder
             * Secure Key
             * 
             *************************************************/
           // string[] Status = _Main.LoadCheck();
            string[] Status = { "true", "true", "true", "true", "true" };
            if (Status.Length == 5)
            {

                // if (Status[4] == "false")
                if (Status[4] == "true")
                {
                    Array.Resize(ref args, 3);



                    args[0] = @"D:\Temp_files\test13\InputPDF";
                    args[1] = @"D:\Temp_files\test13\OutputXLSX";
                    args[2] = @"D:\Temp_files\test13\ProcessedPDF";

                    //args[0] = ConfigurationManager.AppSettings["InputPDF"];
                    //args[1] = ConfigurationManager.AppSettings["OutputExcel"];
                    //args[2] = ConfigurationManager.AppSettings["ProcessedPDF"];



                    //args[0] = @"\\fbd-vs1\PublicNew\RPA_MEDIA\INPUT\OCR\I_FINAL ESTIMATES";
                    //args[1] = @"\\fbd-vs1\PublicNew\RPA_MEDIA\INPUT\RPA\I_FINAL ESTIMATES";
                    //args[2] = @"\\fbd-vs1\PublicNew\RPA_MEDIA\PROCESSED\P_FINALESTIMATES_PDF";

                    //args[]

                    //Array.Resize(ref args, 3);
                    //args[0] = @"\\Aibotsdc\rpa_media\INPUT\OCR\I_FINALESTIMATES";
                    //args[1] = @"\\Aibotsdc\rpa_media\RPA\EXCEL";
                    //args[2] = @"\\Aibotsdc\rpa_media\PROCESSED\P_FINALESTIMATES_PDF";


                    LogHandeler.ModuleLogFile("---", "---", "---", "Program", "Main","Data from :  " + args[0] +" Data to : " + args[1] +"  Data Processed : " + args[2], null, "63", null, null, null);

                    //run for all files
                    int args_Length = args.Length;
                    if (args_Length > 0)
                    {
                        File_Operations fs = new File_Operations();
                        string[] Input_PDFfiles = fs.GetInputFiles(args[0]);
                        LogHandeler.ModuleLogFile("---", "---", "---", "Program", "Main", "Number of files found : " + Input_PDFfiles.Length.ToString(), null, "58", null, null, null);
                        if (Input_PDFfiles.Length > 0)
                        {

                            foreach (string eachpdf in Input_PDFfiles)
                            {
                                try
                                {
                                    releaseExcel();
                                    if (fs.ProcessedOrNot(eachpdf, 1))
                                    { 
                                        Console.WriteLine("Process started " + eachpdf);
                                        //GetFileName to bind output for temp and csv
                                        string fileName = Path.GetFileNameWithoutExtension(eachpdf) + ".xlsx";
                                        string outputExcelfile = Path.Combine(args[1], fileName);

                                        string InternalTempFileLocation = @"C:\ABT_ENV\Helix\Basic";

                                         
                                        LogHandeler.ModuleLogFile(eachpdf, outputExcelfile, TotalPages(eachpdf).ToString(), "Program", "Main", null, null, "73", null, null, null);
                                        //File Operationf for Intrnal pdf.
                                        string ExcelLocation = InternalTempFileLocation + @"\Processed_initial";
                                        bool exists = System.IO.Directory.Exists(ExcelLocation);
                                        if (!exists)
                                        {
                                            System.IO.Directory.CreateDirectory(ExcelLocation);
                                            LogHandeler.ModuleLogFile("---", "---", "---", "Program", "Main", ExcelLocation + "folder not found, New folder created", null, "80", null, null, null);
                                        }
                                        else
                                        {
                                            LogHandeler.ModuleLogFile("---", "---", "---", "Program", "Main", ExcelLocation + "folder found", null, "84", null, null, null);
                                        }
                                        string ExcelFileName = Path.GetFileNameWithoutExtension(eachpdf) + ".xlsx";

                                        string TempoutputExcelfile = Path.Combine(ExcelLocation, ExcelFileName);
                                        LogHandeler.ModuleLogFile("---", "---", "---", "Program", "Main", "Temp Excel File :" + TempoutputExcelfile, null, "89", null, null, null);
                                        if (File.Exists(TempoutputExcelfile))
                                        {
                                            File.Delete(TempoutputExcelfile);
                                            LogHandeler.ModuleLogFile("---", "---", "---", "Program", "Main", "Temp Excel File :" + TempoutputExcelfile + "Found Hence deleted", null, "93", null, null, null);

                                        }


                                        Core_Operation1 av = new Core_Operation1();
                                        //bool processTable_Status = av.ProcsessTable(eachpdf, outputExcelfile, args[2]);
                                        string[,] processTable_Status = av.ProcsessTable(eachpdf, TempoutputExcelfile, args[2]);

                                        if (processTable_Status.GetLength(0) > 0)
                                        {
                                            LogHandeler.ModuleLogFile("---", "---", "---", "Program", "Main", "Wrapper 1 Completed from Core_Operation1", null, "104", null, null, null);
                                            Wrapper wrapper = new Wrapper();
                                            wrapper.MonthNewExcel(TempoutputExcelfile, args[1], processTable_Status);
                                            LogHandeler.ModuleLogFile("---", "---", "---", "Program", "Main", "Wrapper 2 Completed from Wrapper", null, "107", null, null, null);

                                        }
                                        else
                                        {
                                            fs.ProcessedOrNot(eachpdf, 2);
                                            LogHandeler.ModuleLogFile("---", "---", "---", "Program", "Main", "Wrapper 1 Completed from Core_Operation1 but no output returned", null, "112", null, null, null);

                                        }
                                        Console.WriteLine("Processed " + eachpdf);
                                    }
                                    else
                                    {
                                        Console.WriteLine("Exception  : Unable to process the file, Moved to Exception." + eachpdf);
                                    }

                                }
                                catch (Exception ex)
                                {
                                    fs.ProcessedOrNot(eachpdf, 2);
                                    LogHandeler.ModuleLogFile("---", "---", "---", "Program", "Main", "Exception occured", "Yes", "121", ex.StackTrace, ex.ToString(), "Exception");

                                }
                            }
                            goto ProgramStart;
                        }
                        else
                        {
                            goto EndFunction;
                        }


                    }
                }
                else
                {
                    MessageBox.Show(Status[3], "License", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    LogHandeler.ModuleLogFile(Status[3], "License", null, "Program", "Main", "Exception occured", null, "129", null, null, null);

                    try
                    {
                        string location = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                        string file = "ABT_Helix_Basic_SupportFile.exe";
                        // Check if file exists with its full path

                        file = "ABT_Helix_Basic_SupportFile.exe";
                        if (File.Exists(Path.Combine(location, file)))
                        {
                            LogHandeler.ModuleLogFile(Status[3], null, null, "Program", "Main", "file delted from " + location, null, "143", null, null, null);
                            File.Delete(Path.Combine(location, file));
                        }
                    }
                    catch (IOException ex)
                    {
                        LogHandeler.ModuleLogFile("---", "---", "---", "Program", "Main", "Exception occured", "Yes", "121", ex.StackTrace, ex.ToString(), "Exception");

                        // Console.WriteLine(ioExp.Message);
                    }
                }
            }
            else
            {

                LogHandeler.ModuleLogFile("Missing register License. Please contact AiBots Tech", "License", null, "Program", "Main", null, null, "151", null, null, null);

                MessageBox.Show("Missing register License. Please contact AiBots Tech", "License", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        EndFunction:
            releaseExcel();
        }
        static int TotalPages(string PDF)
        {
            LogHandeler.ModuleLogFile("---", "---", "---", "Program", "TotalPages", "Function started", null, "15", null, null, null);

            PdfReader pdfReader = new PdfReader(PDF);
            return pdfReader.NumberOfPages;
        }

        private static void releaseExcel()
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "/C taskkill /IM excel.exe /F";
            startInfo.Verb = "runas";
            process.StartInfo = startInfo;
            process.Start();
            //System.Threading.Thread.Sleep(3000);
            //process.Close();
        }




    }
}
