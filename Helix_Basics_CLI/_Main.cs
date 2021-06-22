using ABT.License;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace Helix_Basics_CLI
{
    static class _Main
    {
        internal static string[] LoadCheck()
        {
            LogHandeler.ModuleLogFile("---", "---", "---", "_Main", "LoadCheck()", "Function started", null, "13", null, null, null);
            string[] info = new string[5];
            info[0] = ComputerInfo.GetComputerId();
            LogHandeler.ModuleLogFile("", "", "", "_Main", "LoadCheck", "Computer ID" + info[0], null, "15", null, null, null);
            //lblProductID.Text = "";// ComputerInfo.GetComputerId();
            KeyManager km = new KeyManager(info[0]);

            LicenseInfo lic = new LicenseInfo();
            //Get license information from license file
            string location = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            //Console.WriteLine(location);
            int value = km.LoadSuretyFile(string.Format(@"{0}\Key.lic", location), ref lic);
            string productKey = lic.ProductKey;
            //Check valid
            if (km.ValidKey(ref productKey))
            {
                KeyValuesClass kv = new KeyValuesClass();
                //Decrypt license key
                if (km.DisassembleKey(productKey, ref kv))
                {
                    info[1] = "ABT Helix Basic";
                    //lblProductKey.Text = productKey;
                    if (kv.Type == LicenseType.EXPIRE)
                    {
                        info[2] = string.Format("{0} days", (kv.Expiration - DateTime.Now.Date).Days);
                    
                    }
                    else
                    {
                        info[2] = "Never_Expire";
                        info[4] = "true";
                    }
                        
                        
                    int datesremaing = Convert.ToInt32((kv.Expiration - DateTime.Now.Date).Days.ToString());
                    if (datesremaing < 0 && datesremaing >= -30 )
                    {

                        LogHandeler.ModuleLogFile("", "", "", "_Main", "LoadCheck", "Product Service expired, Your are in trail period. Product will shutdown in " + (30 + datesremaing).ToString() + " days. Contact AIBOTS", null, "44", null, null, null);
                        info[4] = "true";
                        info[3] = "Product Service expired, Your are in trail period. Product Service will shutdown in " + (30 + datesremaing).ToString() + " days. Contact AIBOTS";
                        MessageBox.Show("Product Service expired, Your are in trail period. Product service will shutdown in " + (30 + datesremaing).ToString() + " days. Contact AIBOTS", "ABT Helix Basic License", MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                    }
                    else if (datesremaing == 0)
                    {
                        info[4] = "true";
                        info[3] = "Product will expire today. Contact AIBOTS";
                        LogHandeler.ModuleLogFile("", "", "", "_Main", "LoadCheck", "Product service will expire today. Contact AIBOTS", null, "53", null, null, null);
                        MessageBox.Show("Product Service will expire today. Contact AIBOTS", "ABT Helix Basic", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else if(datesremaing < -30)
                    {
                        info[4] = "true";
                        info[3] = "Product expired. Contact AIBOTS";
                        LogHandeler.ModuleLogFile("", "", "", "_Main", "LoadCheck", "Product Service expired.Contact AIBOTS", null, "60", null, null, null);
                        MessageBox.Show("Product Service expired. Contact AIBOTS", "ABT Helix Basic", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                    }
                    else
                    {
                        info[4] = "true";
                        info[3] = "Product Service active for " + (datesremaing).ToString() + " days.";
                        LogHandeler.ModuleLogFile("", "", "", "_Main", "LoadCheck", "Product Service active for " + (datesremaing).ToString() + " days.", null, "68", null, null, null);
                    }

                    Support();
                }
            }
            return info;
        }

        static void Support()
        {
            LogHandeler.ModuleLogFile("---", "---", "---", "_Main", "Support()", "Function started", null, "66", null, null, null);
            string location = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            if (!File.Exists(@"c:\ABT_ENV\Helix\Basic\Support_File\ABT_Helix_Basic_SupportFile.exe"))
            {
                LogHandeler.ModuleLogFile("", "", "", "_Main", "Support", "ABT_Helix_Basic_SupportFile.exe not found at" + location, null, "68", null, null, null);
                FileCheck:
                if(Directory.Exists(@"c:\ABT_ENV\Helix\Basic\Support_File"))
                {

                    LogHandeler.ModuleLogFile("", "", "", "_Main", "Support", "ABT_Helix_Basic_SupportFile.exe not found at" + location + "Hence copied form source", null, "73", null, null, null);

                    File.Copy(Path.Combine(location, "ABT_Helix_Basic_SupportFile.exe"), @"c:\ABT_ENV\Helix\Basic\Support_File\ABT_Helix_Basic_SupportFile.exe", true);

                }
                else
                {
                    LogHandeler.ModuleLogFile("", "", "", "_Main", "Support",  location + "Not found. Hence created directory", null, "80", null, null, null);

                    Directory.CreateDirectory(@"c:\ABT_ENV\Helix\Basic\Support_File");
                    goto FileCheck;
                }
                
            }
        }

    } 


    /* log files */
    public static class LogHandeler
    {
        public static void ModuleLogFile(string InputFileName, string OutputfileName, string Totalpages, string ClassName, string FunctionName, string Report, string ExceptionStatus, string LineNumber, string StackTract, string ExceptionMessage, string Status)
        {
            //Date,Time,FileName,OutputFileName,Totalpages,class,function,report,exception Y/N,line number, StackTrace, exception message,OuputStatus

            string LOG = DateTime.Now + "," + InputFileName + "," + OutputfileName + "," + Totalpages + "," + ClassName + "," + FunctionName + "," + Report + "," + ExceptionStatus + "," + LineNumber + "," +"\"" + StackTract + "\"" + "," + "\"" + ExceptionMessage + "\"" + "," + Status;
            //StreamWriter log;
            string LogFile = @"c:\ABT_ENV\Helix\Basic\Log\Module.log";
            if (!Directory.Exists(@"c:\ABT_ENV\Helix\Basic\Log"))
                Directory.CreateDirectory(@"c:\ABT_ENV\Helix\Basic\Log\");
            using (StreamWriter sw = (File.Exists(LogFile)) ? File.AppendText(LogFile) : File.CreateText(LogFile))
            {
                sw.WriteLine(LOG);
                sw.Close();

            }


            //if (!File.Exists(LogFile))
            //{
            //    File.Create(LogFile);
            //    //log = File.AppendText(LogFile);
            //    using (StreamWriter sw = (File.Exists(LogFile)) ? File.AppendText(LogFile) : File.CreateText(LogFile))
            //    {

            //    }
            //}
            //else
            //{
            //    log = File.AppendText(LogFile);
            //}
            //// Write to the file:
            //log.WriteLine(LOG);
            // Close the stream:
            //log.Close();
        }
        public static void ExceptionLogFile(string InputFileName, string OutputfileName, string Totalpages, string ClassName, string FunctionName,string LineNumber, string StackTract, string ExceptionMessage, string Status)
        {
            //Date,Time,FileName,OutputFileName,Totalpages,class,function,report,exception Y/N,line number, StackTrace, exception message,OuputStatus

            string LOG = DateTime.Now + ", " + InputFileName + "," + OutputfileName + "," + Totalpages + "," + ClassName + "," + FunctionName  + "," + LineNumber + "," +"\"" + StackTract + "\"" + "," + "\"" + ExceptionMessage + "\"" + "," + Status;
            StreamWriter log;
            string LogFile = @"c:\ABT_ENV\Helix\Basic\Log\Exception.log";
            if (!File.Exists(LogFile))
            {
                log = new StreamWriter(LogFile);
            }
            else
            {
                log = File.AppendText(LogFile);
            }
            // Write to the file:
            log.WriteLine(LOG);
            // Close the stream:
            log.Close();
        }
        public static void UserLog(string InputFileName, string OutputfileName, string Totalpages, string ClassName, string FunctionName, string Report, string ExceptionStatus, string LineNumber, string StackTract, string ExceptionMessage, string Status)
        {
            //Date,Time,FileName,OutputFileName,Totalpages,class,function,report,exception Y/N,line number, StackTrace, exception message,OuputStatus

            string LOG = DateTime.Now + "," + InputFileName + "," + OutputfileName + "," + Totalpages + "," + ClassName + "," + FunctionName + "," + Report + "," + ExceptionStatus + "," + LineNumber + "," +"\"" + StackTract + "\"" + "," + "\"" + ExceptionMessage + "\"" + "," + Status;
            StreamWriter log;
            string LogFile = @"c:\ABT_ENV\Helix\Basic\Log\Report.log";
            if (!File.Exists(LogFile))
            {
                log = new StreamWriter(LogFile);
            }
            else
            {
                log = File.AppendText(LogFile);
            }
            // Write to the file:
            log.WriteLine(LOG);
            // Close the stream:
            log.Close();
        }


    }
    public static class ExceptionHelper
    {
        public static int LineNumber(this Exception e)
        {
            int linenum = 0;
            try
            {
                linenum = Convert.ToInt32(e.StackTrace.Substring(e.StackTrace.LastIndexOf(":line") + 5));
            }
            catch
            {
                //Stack trace is not available!
            }
            return linenum;
        }
    }


    class ParkingLot
    {
        // initial triggers plan
        internal bool ProcessTable(string inputfile, string outputfile)
        {
            // goto below;

            string InternalFileLocation = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            //outputfile = @"C:\ABT_ENV\Helix\Temp\1.csv";

            var p = new Process();
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            p.StartInfo.RedirectStandardOutput = true;
            string eOut = null;
            p.StartInfo.RedirectStandardError = true;
            p.ErrorDataReceived += new DataReceivedEventHandler((sender, e) => { eOut += e.Data; });
            p.StartInfo.FileName = @"d:\ABT_Helix_Basic_SupportFile.exe";

            p.StartInfo.Arguments = @"-l -i" + inputfile + " -o " + outputfile + " -p all";

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

            p.Start();

            // To avoid deadlocks, use an asynchronous read operation on at least one of the streams.  
            //p.BeginErrorReadLine();
            //string output = p.StandardOutput.ReadToEnd();
            string output = null;
            //p.WaitForExit();
            if (string.IsNullOrEmpty(output) || string.IsNullOrWhiteSpace(output))
            {
                // no data processed
            }
            else
            {
                //Log processed.
            }
        //Console.WriteLine($"The last 50 characters in the output stream are:\n'{output.Substring(output.Length - 50)}'");
        //Console.WriteLine($"\nError stream: {eOut}");
 

            return true;
        }
    }
}
