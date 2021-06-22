using iTextSharp.text.pdf;
using System;
using System.Diagnostics;
using System.IO;
using System.Text;


namespace Helix_Basics_CLI
{
    class File_Operations
    {
        internal string[] GetInputFiles(string Directory_Locaiton)
        {
            try
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "File_Operations", "GetInputFiles", "Function started. From : " + Directory_Locaiton, null, "14", null, null, null);

                string[] filePaths = Directory.GetFiles(@"" + Directory_Locaiton, "*.pdf");
                return filePaths;
            }
            catch (Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "File_Operations", "movefile", "Exceptions Occured during file move during Exceptional PDF", "Yes", "104", ex.StackTrace, ex.ToString(), "Exception");
                return null;
            }
        }

        internal bool ProcessedOrNot(string filename, int i)
        {
            PdfReader reader = null;
            try
            {
                if (i == 2)
                {
                    movefile(filename);
                }
                else
                { 
                    StringBuilder text1 = new StringBuilder();
                    if (File.Exists(filename))
                    {
                        using (reader = new PdfReader(filename))
                        {
                            if (reader.NumberOfPages > 0)
                            {
                                //text1.Append("::::::::::");
                                text1.Append(iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(reader,1));
                                reader.Close();
                                if (text1.ToString().Length > 0)
                                {
                                    if(text1.ToString().Contains("Television Estimate"))
                                    {
                                        //reader.Close();
                                        return true;
                                    }
                                    else
                                    {
                                        //reader.Close();
                                        movefile(filename);
                                    }
                                    
                                }
                                else
                                {
                                    //reader.Close();
                                    movefile(filename);
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                reader.Close();
                movefile(filename);
                return false;

            }
            return false;
        }

        internal void movefile(string file)
        {
            try
            {

                ReleaseMemory("");
                string ProcessedPDFfileName = Path.GetFileNameWithoutExtension(file);
                string DestinationLocation = Path.GetDirectoryName(file);
                DestinationLocation = Path.Combine(DestinationLocation, "Exception");


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
                
                try
                {
                    File.Move(file, DestinatioFileSave);

                   // File.Delete(file);

                }
                catch
                {
                    File.Move(file, DestinatioFileSave);
                    File.Copy(file, DestinatioFileSave,true);
                    ReleaseMemory("");
                    System.Threading.Thread.Sleep(2000);
                    deletefile(file);
                }
                
                LogHandeler.ModuleLogFile("---", "---", "---", "File_Operations", "movefile", "PDF moved for Exception folder " + file + "To: " + DestinatioFileSave, null, "100", null, null, null);
            }
            catch(Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "File_Operations", "movefile", "Exceptions Occured during file move during Exceptional PDF", "Yes", "104", ex.StackTrace, ex.ToString(), "Exception");

            }

        }
        Process process = new System.Diagnostics.Process();
        System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
        private void ReleaseMemory(string processName)
        {
            
           
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            //startInfo.Arguments = "/C taskkill /IM excel.exe /F";
            startInfo.Arguments = "/C taskkill /IM AcroRd32.exe /F";
            startInfo.Verb = "runas";
            process.StartInfo = startInfo;
            process.Start();

            //System.Threading.Thread.Sleep(3000);
            //process.Close();
            
        }

        private void deletefile(string file)
        {


            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            //startInfo.Arguments = "/C taskkill /IM excel.exe /F";
            startInfo.Arguments = "/C del /f " + file;
            startInfo.Verb = "runas";
            process.StartInfo = startInfo;
            process.Start();

            //System.Threading.Thread.Sleep(3000);
            //process.Close();

        }
    }
}
