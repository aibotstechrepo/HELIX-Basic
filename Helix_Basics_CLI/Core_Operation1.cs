using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Helix_Basics_CLI
{
    class Core_Operation1
    {
        public string[,] ProcsessTable(string inputfile, string outputfile, string Processed_Location)
        {
            try
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "ProcsessTable", "Function started. Data:---- inputfile : " + inputfile + "   outputfile  : " + outputfile + "   Processed_Location  : " + Processed_Location, null, "18", null, null, null);
                Core_InnerOperations cIO = new Core_InnerOperations();
                string[] InternalPDF0InternalCSV1 = TempFile(inputfile);
                if (string.IsNullOrEmpty(InternalPDF0InternalCSV1[0]) == false && string.IsNullOrEmpty(InternalPDF0InternalCSV1[1])==false)
                {
                    try
                    { 
                        bool data = cIO.ProcessTable(InternalPDF0InternalCSV1[0], InternalPDF0InternalCSV1[1]);
                        if (data)
                        {
                            try
                            {

                                Core_Operation2 Cop = new Core_Operation2();
                                string[,] TableProcess_status = Cop.TableProcess(InternalPDF0InternalCSV1[0], InternalPDF0InternalCSV1[1], outputfile, inputfile);
                                if (TableProcess_status.GetLength(0) > 0)
                                {
                                    try
                                    { 
                                        string fileName = Path.GetFileName(inputfile);
                                        string fullPath = Path.Combine(Processed_Location, fileName);
                                        bool DeleteTempMoveProcessed_status = DeleteTempMoveProcessed(InternalPDF0InternalCSV1[0], InternalPDF0InternalCSV1[1], inputfile, fullPath);
                                        if (DeleteTempMoveProcessed_status)
                                        {
                                            LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "ProcsessTable", "DeleteTempMoveProcessed_status retuned true Data : ---- InternalPDF0InternalCSV1[0] : "+ InternalPDF0InternalCSV1[0] + "  InternalPDF0InternalCSV1[1] :  "+ InternalPDF0InternalCSV1[1]  + " inputfile :"+ inputfile + "  fullPath"+ fullPath , "no", "42", null, null, null);

                                            return TableProcess_status;
                                        }
                                        else
                                        {
                                            LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "ProcsessTable", " Unable to process DeleteTempMoveProcessed_status retuned false .Data : ---- InternalPDF0InternalCSV1[0] : " + InternalPDF0InternalCSV1[0] + "  InternalPDF0InternalCSV1[1] :  " + InternalPDF0InternalCSV1[1] + " inputfile :" + inputfile + "  fullPath" + fullPath, "no", "42", null, null, null);

                                            return new string[0, 0];
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "ProcsessTable", "Exceptions Occured", "Yes", "55", ex.StackTrace, ex.ToString(), "Exception");
                                        return new string[0, 0];
                                    }

                                }
                                else
                                {
                                    return new string[0, 0];
                                }
                            }
                            catch (Exception ex)
                            {
                                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "ProcsessTable", "Exceptions Occured", "Yes", "67", ex.StackTrace, ex.ToString(), "Exception");
                                return new string[0, 0];
                            }

                        }
                        else
                        {

                            LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "ProcsessTable", "ProcessTable returned false. Data :---- InternalPDF0InternalCSV1[0] :  " + InternalPDF0InternalCSV1[0] + "   InternalPDF0InternalCSV1[1]  :  " + InternalPDF0InternalCSV1[1], "No", "75", null, null, null);
                            return new string[0, 0];
                        }
                    }
                    catch(Exception ex)
                    {
                        LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "ProcsessTable", "Exceptions Occured","Yes", "81",ex.StackTrace, ex.ToString(), "Exception");
                        return new string[0, 0];
                    }
                }
                else
                {
                    LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "ProcsessTable", "TempFile Closed with null values", null, "55", null, null, null);
                    return new string[0, 0];
                }
               
            }
            catch(Exception ex)
            {//call LogFile method and pass argument as Exception message, event name, control name, error line number, current form name
             // LogHandeler.LogFile(exe.Message, "Core_Operation1->ProcsessTable", "","ProcsessTable", exe.LineNumber(), exe.Data.ToString());
             // return false;

                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "ProcsessTable", "Exceptions Occured", "Yes", "97", ex.StackTrace, ex.ToString(), "Exception");
                return new string[0, 0];
            } 
        }
        private string[] TempFile(string SourcePDFFile)
        {
            LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "TempFile", "Function started. Source PDF:-- "+ SourcePDFFile, null, "103", null, null, null);
            string[] InternalPDF0InternalCSV1 = new string[3];
            InternalPDF0InternalCSV1[0] = null; //PDF
            InternalPDF0InternalCSV1[1] = null; //CSV
            InternalPDF0InternalCSV1[2] = null; //Exception
            try
            {
                string InternalTempFileLocation = @"C:\ABT_ENV\Helix\Basic";
                
                //File Operationf for Intrnal pdf.
                string PdfLocation = InternalTempFileLocation + @"\InputFiles";
                bool exists = System.IO.Directory.Exists(PdfLocation);
                if (!exists)
                    System.IO.Directory.CreateDirectory(PdfLocation);

                
                string PDFfileName = Path.GetFileName(SourcePDFFile);
                string PDFfullPath = Path.Combine(PdfLocation, PDFfileName);
                if (File.Exists(PDFfullPath))
                {
                    File.Delete(PDFfullPath);
                    LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "TempFile", "TEMP PDF exists, Deleted file from " + PDFfullPath, null, "124", null, null, null);
                } 

                File.Copy(SourcePDFFile, PDFfullPath, true);
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "TempFile", "copied source pdf to temp. Source: " + SourcePDFFile +" Destination: " + PDFfullPath, null, "128", null, null, null);
            
                InternalPDF0InternalCSV1[0] = PDFfullPath;


                //File Opertaion for Internal csv.
                //string InternalFileLocation = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

                string CSVLocation = InternalTempFileLocation + @"\TempFile";
                exists = System.IO.Directory.Exists(CSVLocation);
                if (!exists)
                    System.IO.Directory.CreateDirectory(CSVLocation);
                string fileName = Path.GetFileNameWithoutExtension(SourcePDFFile) +".csv";
                string fullPath = Path.Combine(CSVLocation, fileName);

                if (File.Exists(fullPath))
                {
                    File.Delete(fullPath);
                    LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "TempFile", "Destination file exists hence deleted the previous file from "+ fullPath, null, "146", null, null, null);
                }


                InternalPDF0InternalCSV1[1] = fullPath;



            }
            catch(Exception ex)
            {
                InternalPDF0InternalCSV1[2] = ex.ToString(); 
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "TempFile", "Exceptions Occured", "Yes", "158", ex.StackTrace, ex.ToString(), "Exception");

            }
            return InternalPDF0InternalCSV1;
        }
         

        private bool DeleteTempMoveProcessed(string InternalPDF, string TempFile, string FileFrom, string FileTo)
        {
            LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "DeleteTempMoveProcessed", "Function started", null, "167", null, null, null);

            try
            {
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
                File.Delete(TempFile);
                File.Delete(InternalPDF);
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "DeleteTempMoveProcessed", "deleted temp file from " + TempFile, null, "173", null, null, null);
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "DeleteTempMoveProcessed", "deleted temp file from " + InternalPDF, null, "174", null, null, null);

            }
            catch (Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "DeleteTempMoveProcessed", "Exceptions Occured", "Yes", "179", ex.StackTrace, ex.ToString(), "Exception");
            }
            try
            {

                string ProcessedPDFfileName = Path.GetFileNameWithoutExtension(FileTo);
                string DestinationLocation = Path.GetDirectoryName(FileTo);



                string DestinatioFileSave = Path.Combine(DestinationLocation, ProcessedPDFfileName);
                renamepdf:
                if (File.Exists(DestinatioFileSave + ".pdf"))
                {
                    LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "DeleteTempMoveProcessed", "Found Duplicate entry. New File Name" + DestinatioFileSave + ".pdf", null, "193", null, null, null);
                    DestinationLocation = Path.Combine(DestinationLocation, "Duplicate");
                    DestinatioFileSave = Path.Combine(DestinationLocation, ProcessedPDFfileName);
                    if (File.Exists(DestinatioFileSave + ".pdf"))
                    {
                        string Time = DateTime.Now.ToString();

                        Time = Time.Replace(" ", "_");
                        Time = Time.Replace(":", "_");
                        DestinatioFileSave = Path.Combine(DestinationLocation, ProcessedPDFfileName + "_" + Time);
                    }
                }
                DestinatioFileSave = DestinatioFileSave + ".pdf";
                bool folderExists = Directory.Exists(DestinationLocation);
                if (!folderExists)
                    Directory.CreateDirectory(DestinationLocation); 
                if (File.Exists(DestinatioFileSave))
                {
                    goto renamepdf;
                }
                File.Move(FileFrom, DestinatioFileSave);
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "DeleteTempMoveProcessed", "PDF moved from " + FileFrom + "To: " + DestinatioFileSave, null, "214", null, null, null);

                // return true;
            }
            catch (Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation1", "DeleteTempMoveProcessed", "Exceptions Occured", "Yes", "220", ex.StackTrace, ex.ToString(), "Exception");
            }
            return true;
        } 
    }
    class Core_InnerOperations
    {
        internal bool ProcessTable(string inputfile, string outputfile)
        {
            try
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_InnerOperations", "ProcessTable", "Function started. (cdms)Data:- inputfile  :  " + inputfile + "   outputfile  : " + outputfile, null, "233", null, null, null);

                string InternalFileLocation = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                //Console.WriteLine("project run location : " +InternalFileLocation);
                //outputfile = @"C:\ABT_ENV\Helix\Temp\1.csv";

                //inputfile = "d:\\11.pdf";


                string command = @"c:\ABT_ENV\Helix\Basic\Support_File\ABT_Helix_Basic_SupportFile.exe -l " + inputfile + " -o " + outputfile + " -p all -r";
                //Console.WriteLine(command);
                bool check = RunCommand(command, outputfile);
                if (check)
                    return true;
                #region info
                /* param:  
                 * """""ABT_Helix_Basic_SupportFile.exe -l D:\1.pdf -o D:\Temp_files\test2\testingcs.csv -p all"""""
                 * ABT_Helix_Basic_SupportFile.exe :- our phase 1 exe
                 * -l : https://github.com/tabulapdf/tabula-java l,--lattice               Force PDF to be extracted using lattice-mode
                                                                                            extraction (if there are ruling lines
                                                                                            separating each cell, as in a PDF of an Excel
                                                                                            spreadsheet)
                 * D:\1.pdf : Source PDF
                 * -o :  Write output to <file> instead of STDOUT default -
                 * D:\Temp_files\test2\testingcs.csv: Output location
                 * -p,--pages <PAGES>                                                      Comma separated list of ranges, or all.
                                                                                            Examples: --pages 1-3,5-7, --pages 3 or
                                                                                            --pages all. Default is --pages 1
                 * 

                 */
                #endregion
            }
            catch (Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_InnerOperations", "ProcessTable", "Exceptions Occured", "Yes", "263", ex.StackTrace, ex.ToString(), "Exception");
            }
            // goto below; 
            return false;
        }
        Process process = new Process();
        bool RunCommand(string command, string outFile)
        {
            try
            {

                LogHandeler.ModuleLogFile("---", "---", "---", "Core_InnerOperations", "runCommand", "Function started. Data: C -"  + command + "   F  - " + outFile, null, "277", null, null, null);

                //* Create your Process

                process.StartInfo.FileName = "cmd.exe";
                //string command = @"d:\ABT_Helix_Basic_SupportFile.exe -l D:\1.pdf -o D:\Temp_files\test13\asdf.csv -p all";
                process.StartInfo.Arguments = "/c " + command;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                //* Set your output and error (asynchronous) handlers
                process.OutputDataReceived += new DataReceivedEventHandler(OutputHandler);
                process.ErrorDataReceived += new DataReceivedEventHandler(OutputHandler);
                //* Start process and handlers
                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                //if(Console_Output == "error")
                //{
                //    process.Close();
                //}
                process.WaitForExit();
                bool checksize = checktoprocessed(outFile);
                if (checksize)
                    return true;
                return false;
            }
            catch (Exception ex)
            {

                LogHandeler.ModuleLogFile("---", "---", "---", "Core_InnerOperations", "runCommand", "Run Command: " + command + "Exceptions Occured", "Yes", "263", ex.StackTrace, ex.ToString(), "Exception");

            }
            return false;

        }

        void OutputHandler(object sendingProcess, DataReceivedEventArgs outLine)
        {
            //* Do your stuff with the output (write to console/log/StringBuilder)
            string data = outLine.Data;
            if (string.IsNullOrEmpty(data) == false)
            {
                if (data.ToLower().Trim() == "spreadsheet")
                {
                    LogHandeler.ModuleLogFile("---", "---", "---", "Core_InnerOperations", "OutputHandler", "Output String: " + data, null, "316", null, null, null);
                    //working
                    //Console.WriteLine(outLine.Data);
                }
                else
                {
                    //not working, exceptions

                    LogHandeler.ModuleLogFile("---", "---", "---", "Core_InnerOperations", "OutputHandler", "Process Stoped: " + data, null, "324", null, null, null);
                   // LogHandeler.ModuleLogFile("---", "---", "---", "Core_InnerOperations", "OutputHandler", "Process stoped", null, "325", null, null, null);

                    //process.Close();
                    //Console_Output = "error";
                }
            }
        }

        bool checktoprocessed(string outFile)
        {
            FileInfo fi = new FileInfo(outFile);
            if (fi.Exists)
            {
                long size = fi.Length;
                if(size > 20)
                {
                    return true;
                }
            }
            return false;
        }
    }

    

}
