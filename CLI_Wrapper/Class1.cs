using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Reflection;

namespace CLI_Wrapper
{
    public class Wrapper
    { 
        public bool MonthNewExcel(string SourceFile, string DestinationFolder, string[,] month_year)
        {
            LogHandeler.ModuleLogFile("---", "---", "---", "Wrapper", "MonthNewExcel", "Function started", null, "12", null, null, null);

            bool status = false; 
            //Core Function
            //string SourceFile = @"D:\Temp_files\test13\OutputXLSX\2.xlsx";
            //string DestinationFile = @"D:\Temp_files\test13\OutputXLSX\2 - Copy.xlsx";
            //int sheetNumber = 1;

            ApplicationClass excelApplicationClass = new ApplicationClass();
            _Workbook destinationWorkbook = null;
            Workbook sourceWorkbook = null;
            Worksheet sourceworkSheet = null;
            Worksheet destinationWorksheet = null;

            try
            {

                string[] _month = Split2dArray1stColumn(month_year);
                //Source
                excelApplicationClass.Visible = false;
                sourceWorkbook = excelApplicationClass.Workbooks.Open(SourceFile, false, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                int TotalSheets = sourceWorkbook.Worksheets.Count;
                for (int sheetNum = 1; sheetNum < TotalSheets; sheetNum++)
                {
                    try
                    {
                        destinationWorkbook = excelApplicationClass.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                        //Open the WorkSheet
                        sourceworkSheet = (Worksheet)sourceWorkbook.Sheets[sheetNum];
                        int countWorkSheet = destinationWorkbook.Worksheets.Count;
                        destinationWorksheet = (Worksheet)destinationWorkbook.Sheets[countWorkSheet];
                        string SheetName = sourceworkSheet.Name;
                        sourceworkSheet.Copy(Missing.Value, destinationWorksheet);   //Copy from source to destination 

                        sourceworkSheet = (Worksheet)sourceWorkbook.Sheets[sourceWorkbook.Worksheets.Count]; //for Overview
                        countWorkSheet = destinationWorkbook.Worksheets.Count;
                        destinationWorksheet = (Worksheet)destinationWorkbook.Sheets[countWorkSheet];
                        sourceworkSheet.Copy(Missing.Value, destinationWorksheet);

                        

                        destinationWorksheet = (Worksheet)destinationWorkbook.Sheets[1];
                        destinationWorksheet.Delete();

                        destinationWorksheet = (Worksheet)destinationWorkbook.Sheets[1];
                        destinationWorksheet.Name = "Data";
                        int month_Index = Array.FindIndex(_month, element => element == SheetName);
                        if (SheetName.Trim().ToUpper() == "TAXES")
                        {
                            string DestinationFolder_withMonth = Path.Combine(DestinationFolder, month_year[1, _month .Length- 1], "TAXES");
                            string FileName = Path.GetFileNameWithoutExtension(SourceFile) + "_" + SheetName;

                            string DestinatioFileSave = Path.Combine(DestinationFolder_withMonth, FileName);

                            if (File.Exists(DestinatioFileSave + ".xlsx"))
                            {
                                DestinationFolder_withMonth = Path.Combine(DestinationFolder_withMonth, "Duplicate");
                                DestinatioFileSave = Path.Combine(DestinationFolder_withMonth, FileName);
                                if (File.Exists(DestinatioFileSave + ".xlsx"))
                                {
                                    string Time = DateTime.Now.ToString();

                                    Time = Time.Replace(" ", "_");
                                    Time = Time.Replace(":", "_");
                                    DestinatioFileSave = Path.Combine(DestinationFolder_withMonth, FileName+"_"+Time);
                                }
                            }
                            DestinatioFileSave = DestinatioFileSave + ".xlsx";
                            bool folderExists = Directory.Exists(DestinationFolder_withMonth);
                            if (!folderExists)
                                Directory.CreateDirectory(DestinationFolder_withMonth);

                            destinationWorkbook.SaveAs(DestinatioFileSave);

                        }
                        else
                        {

                            string DestinationFolder_withMonth = Path.Combine(DestinationFolder, month_year[1, month_Index], month_year[0, month_Index]);
                            //string FileName = Path.GetFileNameWithoutExtension(SourceFile) + "_" + SheetName + ".xlsx";
                            string FileName = Path.GetFileNameWithoutExtension(SourceFile) + "_" + SheetName;

                            string DestinatioFileSave = Path.Combine(DestinationFolder_withMonth, FileName);

                            if (File.Exists(DestinatioFileSave + ".xlsx"))
                            {
                                DestinationFolder_withMonth = Path.Combine(DestinationFolder_withMonth, "Duplicate");
                                DestinatioFileSave = Path.Combine(DestinationFolder_withMonth, FileName);
                                if (File.Exists(DestinatioFileSave + ".xlsx"))
                                {
                                    string Time = DateTime.Now.ToString(); 

                                    Time = Time.Replace(" ", "_");
                                    Time = Time.Replace(":", "_");
                                    DestinatioFileSave = Path.Combine(DestinationFolder_withMonth, FileName+"_"+ Time);
                                } 
                            }
                            DestinatioFileSave = DestinatioFileSave + ".xlsx";
                            bool folderExists = Directory.Exists(DestinationFolder_withMonth);
                            if (!folderExists)
                                Directory.CreateDirectory(DestinationFolder_withMonth);
                            
                            destinationWorkbook.SaveAs(DestinatioFileSave); 
                        }
                        //Response.Write("Merged !");

                    }
                    catch (Exception ex)
                    {
                        LogHandeler.ModuleLogFile("---", "---", "---", "Wrapper", "MonthNewExcel", "Exceptions Occured", "Yes", "122", "\""+ ex.StackTrace, ex.ToString(), "Exception");
                    }
                    finally
                    {
                        if (destinationWorkbook != null)
                        {
                            destinationWorkbook.Close(true, Missing.Value, Missing.Value);
                        }

                        if (destinationWorkbook != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(destinationWorkbook);

                        destinationWorkbook = null;
                        destinationWorksheet = null;
                    }


                }
                sourceWorkbook.Save();
                sourceWorkbook.Close();
                status = true;
                DelteTempExcel(SourceFile);
            }
            catch (Exception ex)
            {
                status = false;
                LogHandeler.ModuleLogFile("---", "---", "---", "Wrapper", "MonthNewExcel", "Exceptions Occured", "Yes", "148", ex.StackTrace, ex.ToString(), "Exception");
            }
            finally
            {

                //if (sourceWorkbook != null)
                //    sourceWorkbook.Close(true, Missing.Value, Missing.Value);

                //if (sourceworkSheet != null)
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceworkSheet);

                //sourceworkSheet = null;

                //if (sourceWorkbook != null)
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceWorkbook);
                //sourceWorkbook = null;

                if (excelApplicationClass != null)
                {
                    excelApplicationClass.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplicationClass);
                    excelApplicationClass = null;
                } 
            } 
            return status;
        }
        private bool DelteTempExcel(string TempExcelFile)
        {
            LogHandeler.ModuleLogFile("---", "---", "---", "Wrapper", "DelteTempExcel", "Function started", null, "176", null, null, null);

            try
            {
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
                File.Delete(TempExcelFile);
                return true;
            }
            catch(Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Wrapper", "MonthNewExcel", "Exceptions Occured", "Yes", "187", ex.StackTrace, ex.ToString(), "Exception");
                return false;
            }
            
        }
        private string[] Split2dArray1stColumn(string[,] twoDArray)
        {
            string[] data = new string[twoDArray.GetLength(1)];
            for(int i=0; i<=twoDArray.GetLength(1)-1; i++)
            {
                data[i] = twoDArray[0, i];
            }
            return data;
        }
    }
    public static class LogHandeler
    {
        public static void ModuleLogFile(string InputFileName, string OutputfileName, string Totalpages, string ClassName, string FunctionName, string Report, string ExceptionStatus, string LineNumber, string StackTract, string ExceptionMessage, string Status)
        {
            string LOG = DateTime.Now + ", " + InputFileName + "," + OutputfileName + "," + Totalpages + "," + ClassName + "," + FunctionName + "," + LineNumber + "," +"\"" + StackTract + "\"" + "," + "\"" + ExceptionMessage + "\"" + "," + Status;
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
        public static void ExceptionLogFile(string InputFileName, string OutputfileName, string Totalpages, string ClassName, string FunctionName, string LineNumber, string StackTract, string ExceptionMessage, string Status)
        {
            //Date,Time,FileName,OutputFileName,Totalpages,class,function,report,exception Y/N,line number, StackTrace, exception message,OuputStatus

            string LOG = DateTime.Now + "," + InputFileName + "," + OutputfileName + "," + Totalpages + "," + ClassName + "," + FunctionName + "," + LineNumber + "," +"\"" + StackTract + "\"" + "," + "\"" + ExceptionMessage + "\"" + "," + Status;
            //string LOG = DateTime.Now + "," + InputFileName + "," + OutputfileName + "," + Totalpages + "," + ClassName + "," + FunctionName + "," + LineNumber + "," +"\"" + StackTract + "\"" + "," + "\"" + ExceptionMessage + "\"" + "," + Status;
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
}
    
