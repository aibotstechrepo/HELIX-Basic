    using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Reflection;
using System.Text;
using _Excel = Microsoft.Office.Interop.Excel;
namespace Helix_Basics_CLI
{
    class Core_Operation2
    {
        #region Global Variables
        // Default Variables
        private _Application excelCSV = new _Excel.Application();
        private Workbook wbCSV;
        private Worksheet wsCSV;

        private _Application excelEXCEL = new _Excel.Application();
        private Workbook wbEXCEL;
        private Worksheet wsEXCEL; 

        string path = "";

        int tempLastusedrow = 0;

        #endregion
        /*  Main function for data extraction from table.
         *  :- colelctive functions
         */

        internal string[,] TableProcess(string PdfFile, string fileName, string OutfileName, string ActualPDFfromSource)
        {
            try
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "TableProcess", "Function started. Data: PdfFile : " + PdfFile + "  fileName  : "+ fileName + "  OutfileName : " + OutfileName + "  ActualPDFfromSource : " + ActualPDFfromSource, null, "38", null, null, null);

                bool Excel_Create_status = NewExcel(); //Create New Excel to write the tables info

                if (Excel_Create_status)
                {
                    path = OutfileName;
                    bool Open_Excel_Status = Open_Excel(fileName);
                    if (Open_Excel_Status)
                    {
                        int lastRowCustom = wsCSV.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                        if(wsCSV.Cells[lastRowCustom,1].Value2 =="SGS")
                        {
                            wsCSV.Cells[lastRowCustom, 1].Value2 = "SGST";
                            wsCSV.Cells[lastRowCustom-1, 1].Value2 = "CGST";
                            wsCSV.Cells[lastRowCustom-2, 1].Value2 = "TAXES";
                            wbCSV.Save();

                        }
                        int[] loc = FindAll("Program", "Channel", ActualPDFfromSource);
                         

                        //Monday corrections

                        Data_Correction_MON(loc); 
                        if (loc.Length > 0) 
                        {
                            bool MoveTablesMonthWise_staus = MoveTablesMonthWise(loc);
                            if (MoveTablesMonthWise_staus)
                            {

                                string[,] DataController_status = DataController(PdfFile);

                                if (DataController_status.GetLength(0) > 0)
                                {
                                    ReleaseMemory();
                                    return DataController_status;
                                }
                                else
                                {
                                    ReleaseMemory();
                                    return new string[0, 0];
                                }
                            }
                            else
                            {
                                ReleaseMemory();
                                return new string[0, 0];
                            }
                        }
                        else
                        {
                            // ReleaseMemory();
                            return new string[0, 0];
                        }
                    }
                    else
                    {
                        ReleaseMemory();
                        return new string[0, 0];
                    }

                }
                else
                {

                    ReleaseMemory();

                    return new string[0, 0];
                    //unable to create excel - log
                }

            }
            catch(Exception ex)
            { 
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "TableProcess", "Exceptions Occured", "Yes", "104", ex.StackTrace, ex.ToString(), "Exception");
                ReleaseMemory();
                return new string[0, 0];
            }
        }
         
        /* Core Functions:
         * ------------------------
         * Don't modify any codes due to lost of dependesices
         */

        #region Phase 1 - Table processing
        private bool NewExcel() //create new Excel file
        {

            LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "NewExcel", "Function started", null, "117", null, null, null);
            try
            {
                wbEXCEL = excelEXCEL.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                wsEXCEL = wbEXCEL.Worksheets[1];
                excelEXCEL.Visible = true;
                return true;
                //regular log file data
            }
            catch(Exception ex)
            {
 
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "NewExcel", "Exceptions Occured", "Yes", "129", ex.StackTrace, ex.ToString(), "Exception");
                return false;
            }
        }
        private bool Open_Excel(string Path)
        { 
            LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "Open_Excel", "Function started. Path: " + Path, null, "136", null, null, null);
            try
            {
                wbCSV = excelCSV.Workbooks.Open(Path);
                wsCSV = wbCSV.Worksheets[1];
                excelCSV.Visible = false;
                return true;
                //regular log file data
            }
            catch (Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "Open_Excel", "Exceptions Occured. Path : " + Path, "Yes", "147", ex.StackTrace, ex.ToString(), "Exception");
                return false;
            }
        }//Open CSV of Table extraction
        private bool MoveTablesMonthWise(int[] Loc_Data)// Move table data as monthwise
        {
            try
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "MoveTablesMonthWise", "Function started", null, "136", null, null, null);

                
                // Number of Tables to Move - LOG
                for (int i = 0; i < Loc_Data.Length - 1; i++)
                {
                    if (i == Loc_Data.Length - 2)
                    {
                        string temp1 = wsCSV.Cells[Loc_Data[i], 1].Value2.ToString().ToUpper();
                        if ((temp1 == "TAXES")) // if Taxes table found
                        {
                            //Taxes data(GST) - LOG
                            wbEXCEL.Sheets.Add(After: wbEXCEL.Sheets[wbEXCEL.Sheets.Count]);
                            wsEXCEL = wbEXCEL.Worksheets[wbEXCEL.Sheets.Count];
                            wsEXCEL.Name = "TAXES";
                            wsEXCEL.Activate();

                            Range last = wsEXCEL.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                            Range range = wsEXCEL.get_Range("A1", last);


                            int a = GetLastColumn(Loc_Data[i]);
                            

                            int lastUsedRow = last.Row;
                            
                            if (lastUsedRow != 1)
                            {
                                lastUsedRow = lastUsedRow + 2 - tempLastusedrow;
                                tempLastusedrow = 0;
                            }
                             
                            wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow, 1], wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + 1] - Loc_Data[i]) - 1, a]].Value2 = wsCSV.Range[wsCSV.Cells[Loc_Data[i], 1], wsCSV.Cells[Loc_Data[i + 1] - 1, a]].Value2;
                            wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow + 1, 2], wsEXCEL.Cells[lastUsedRow + 1, 2]].NumberFormat = "###,##%";
                            wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow + 2, 2], wsEXCEL.Cells[lastUsedRow + 2, 2]].NumberFormat = "###,##%";

                            wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow + 1, 4], wsEXCEL.Cells[lastUsedRow + 1, 4]].NumberFormat = "###,##%";
                            wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow + 2, 4], wsEXCEL.Cells[lastUsedRow + 1, 4]].NumberFormat = "###,##%";

                            //last = Loc_Data[i + 1] - Loc_Data[i];
                            wsEXCEL.Cells[lastUsedRow, 1].EntireRow.Font.Bold = true;

                            wsEXCEL.Columns.AutoFit();
                            BorderAllSide(lastUsedRow, 1, lastUsedRow + (Loc_Data[i + 1] - Loc_Data[i]) - 1, a);

                        }
                    }


                    string temp = wsCSV.Cells[Loc_Data[i + 1] - 1, 2].Value2.ToString().ToUpper();
                    if ((temp == "TOTAL"))
                    {
                        string CurrentSheet = wsEXCEL.Name.ToString();
                        string[] SheetAvaialble = ExcelSheetNames();

                        if (Array.Exists(SheetAvaialble, element => element == wsCSV.Cells[Loc_Data[i + 1] - 1, 1].Value2.ToString().ToUpper()))
                        {
                            wsEXCEL = wbEXCEL.Worksheets[wsCSV.Cells[Loc_Data[i + 1] - 1, 1].Value2.ToString().ToUpper()];
                            if (CurrentSheet == wsEXCEL.Name.ToString())
                            {
                                ;
                            }
                            else
                            {
                                tempLastusedrow = 0;
                            }
                            wsEXCEL.Activate();
                        }
                        else
                        {
                            wbEXCEL.Sheets.Add(After: wbEXCEL.Sheets[wbEXCEL.Sheets.Count]);
                            wsEXCEL = wbEXCEL.Worksheets[wbEXCEL.Sheets.Count];
                            wsEXCEL.Name = wsCSV.Cells[Loc_Data[i + 1] - 1, 1].Value2.ToString().ToUpper();
                            if (CurrentSheet == wsEXCEL.Name.ToString())
                            {
                                ;
                            }
                            else
                            {
                                tempLastusedrow = 0;
                            }
                            wsEXCEL.Activate();
                        }

                        Range last = wsEXCEL.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                        Range range = wsEXCEL.get_Range("A1", last);

                        int lastUsedRow = last.Row;

                        if (lastUsedRow != 1)
                        { 
                            string tesster = (wsEXCEL.Cells[lastUsedRow, 1].Value2 + ".a").ToString();
                            if (tesster.Length <= 2)
                            {
                                int j = lastUsedRow - 1;
                            dataChecker:
                                tesster = (wsEXCEL.Cells[j, 1].Value2 + ".a").ToString();
                                if (tesster.Length <= 2)
                                {
                                    j--;
                                    goto dataChecker;
                                }
                                else
                                {
                                    lastUsedRow = j;
                                    tempLastusedrow = 0;
                                }
                            }
                        }

                        int a = GetLastColumn(Loc_Data[i]);

                        if (a < 40)
                        {
                            for (int eachCurrentColumn = 1; eachCurrentColumn < 50; eachCurrentColumn++)
                            {
                                if (i< Loc_Data.Length)
                                { 
                                    string cellData = $"{""}{ wsCSV.Cells[Loc_Data[i], eachCurrentColumn].Value2}";
                                    if (string.IsNullOrWhiteSpace(cellData))
                                    {
                                        for (int eachRowInCurrentColumn = Loc_Data[i]; eachRowInCurrentColumn < Loc_Data[i + 1]; eachRowInCurrentColumn++)
                                        {

                                            string cellDatainEveryRow = $"{""}{ wsCSV.Cells[eachRowInCurrentColumn, eachCurrentColumn].Value2}";
                                            if (!string.IsNullOrWhiteSpace(cellDatainEveryRow))
                                            {
                                                //MessageBox.Show("Data found in between");
                                                Range rg1 = (Range)wsCSV.Range[wsCSV.Cells[eachRowInCurrentColumn, eachCurrentColumn - 1], wsCSV.Cells[eachRowInCurrentColumn, eachCurrentColumn - 1]];
                                                rg1.Delete(Shift: XlDeleteShiftDirection.xlShiftToLeft);
                                            }
                                        }
                                        Range rg = (Range)wsCSV.Range[wsCSV.Cells[Loc_Data[i], eachCurrentColumn], wsCSV.Cells[Loc_Data[i + 1] - 1, eachCurrentColumn]];
                                        rg.Delete(Shift: XlDeleteShiftDirection.xlShiftToLeft);

                                    }
                                    else if (cellData.Contains("Net Spot"))
                                    {
                                        wsCSV.Cells[Loc_Data[i], eachCurrentColumn - 1].Value2 = "TotalFCT";
                                    }
                                    else if (cellData.Contains("TUE"))
                                    {
                                        string data = ("a" + wsCSV.Cells[Loc_Data[i], eachCurrentColumn - 1].Value2).ToString().ToUpper().Substring(("a" + wsCSV.Cells[Loc_Data[i], eachCurrentColumn - 1].Value2).ToString().ToUpper().IndexOfAny("0123456789".ToCharArray()));
                                        wsCSV.Cells[Loc_Data[i], eachCurrentColumn - 1].Value2 = "MON" + data;
                                    }
                                }
                                else
                                {
                                    break;
                                }
                            }
                            wbCSV.Save();
                            a = GetLastColumn(Loc_Data[i]);
                        } 
                        if (lastUsedRow != 1)
                        {
                            lastUsedRow = lastUsedRow + 2 - tempLastusedrow;
                            tempLastusedrow = 0;
                        }
                        wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow, 1], wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + 1] - Loc_Data[i]) - 1, a]].Value2 = wsCSV.Range[wsCSV.Cells[Loc_Data[i], 1], wsCSV.Cells[Loc_Data[i + 1] - 1, a]].Value2;
                        //last = Loc_Data[i + 1] - Loc_Data[i];
                        wsEXCEL.Cells[lastUsedRow, 1].EntireRow.Font.Bold = true;
                        string cellValueCheck = (wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + 1] - Loc_Data[i]) - 1, 2].Value2 + ".a").ToString();
                        if (cellValueCheck == "Total.a")
                        {
                            wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + 1] - Loc_Data[i]) - 1, 2].EntireRow.Font.Bold = true;
                        }
                        wsEXCEL.Columns.AutoFit();
                        BorderAllSide(lastUsedRow, 1, lastUsedRow + (Loc_Data[i + 1] - Loc_Data[i]) - 1, a);
                    }
                    else
                    {
                        int a = GetLastColumn(Loc_Data[i+1]);

                        if (a < 41)
                        {
                            for (int eachCurrentColumn = 1; eachCurrentColumn < 50; eachCurrentColumn++)
                            {
                                if (i < Loc_Data.Length)
                                {
                                    string cellData = $"{""}{ wsCSV.Cells[Loc_Data[i + 1], eachCurrentColumn].Value2}";
                                    if (string.IsNullOrWhiteSpace(cellData))
                                    {
                                        if ((i + 2) < Loc_Data.Length)
                                        {
                                            for (int eachRowInCurrentColumn = Loc_Data[i + 1]; eachRowInCurrentColumn < Loc_Data[i + 2]; eachRowInCurrentColumn++)
                                            {
                                                string cellDatainEveryRow = $"{""}{ wsCSV.Cells[eachRowInCurrentColumn, eachCurrentColumn].Value2}";
                                                if (!string.IsNullOrWhiteSpace(cellDatainEveryRow))
                                                {
                                                    //MessageBox.Show("Data found in between");
                                                    Range rg1 = (Range)wsCSV.Range[wsCSV.Cells[eachRowInCurrentColumn, eachCurrentColumn - 1], wsCSV.Cells[eachRowInCurrentColumn, eachCurrentColumn - 1]];
                                                    rg1.Delete(Shift: XlDeleteShiftDirection.xlShiftToLeft);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            break;
                                        }
                                        Range rg = (Range)wsCSV.Range[wsCSV.Cells[Loc_Data[i + 1], eachCurrentColumn], wsCSV.Cells[Loc_Data[i + 2] - 1, eachCurrentColumn]];
                                        rg.Delete(Shift: XlDeleteShiftDirection.xlShiftToLeft);
                                    }
                                    else if(cellData.Contains("Net Spot"))
                                    {
                                        wsCSV.Cells[Loc_Data[i + 1], eachCurrentColumn - 1].Value2 = "TotalFCT";
                                    }
                                    else if(cellData.Contains("TUE"))
                                    { 
                                        string data = ("a" + wsCSV.Cells[Loc_Data[i + 1], eachCurrentColumn - 1].Value2).ToString().ToUpper().Substring(("a" + wsCSV.Cells[Loc_Data[i + 1], eachCurrentColumn - 1].Value2).ToString().ToUpper().IndexOfAny("0123456789".ToCharArray()));
                                        wsCSV.Cells[Loc_Data[i + 1], eachCurrentColumn - 1].Value2 = "MON" + data;
                                    }
                                }
                                else
                                {
                                    break;
                                }
                            }
                            wbCSV.Save();
                            a = GetLastColumn(Loc_Data[i]);
                        }
                        bool flag = true;
                        int j = 2;
                        do
                        {
                            if ((i + j) < Loc_Data.Length - 1)
                            {
                                temp = wsCSV.Cells[Loc_Data[i + j] - 1, 2].Value2.ToString().ToUpper();
                                if ((temp == "TOTAL"))
                                {
                                    string CurrentSheet = wsEXCEL.Name.ToString();
                                    string[] SheetAvaialble = ExcelSheetNames();

                                    if (Array.Exists(SheetAvaialble, element => element == wsCSV.Cells[Loc_Data[i + j] - 1, 1].Value2.ToString().ToUpper()))
                                    {
                                        wsEXCEL = wbEXCEL.Worksheets[wsCSV.Cells[Loc_Data[i + j] - 1, 1].Value2.ToString().ToUpper()];

                                        if (CurrentSheet == wsEXCEL.Name.ToString())
                                        {
                                            ;
                                        }
                                        else
                                        {
                                            tempLastusedrow = 0;
                                        }

                                        wsEXCEL.Activate();
                                    }
                                    else
                                    {
                                        wbEXCEL.Sheets.Add(After: wbEXCEL.Sheets[wbEXCEL.Sheets.Count]);
                                        wsEXCEL = wbEXCEL.Worksheets[wbEXCEL.Sheets.Count];
                                        wsEXCEL.Name = wsCSV.Cells[Loc_Data[i + j] - 1, 1].Value2.ToString().ToUpper();
                                        if (CurrentSheet == wsEXCEL.Name.ToString())
                                        {
                                            ;
                                        }
                                        else
                                        {
                                            tempLastusedrow = 0;
                                        }
                                        wsEXCEL.Activate();
                                    }

                                    Range last = wsEXCEL.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                                    Range range = wsEXCEL.get_Range("A1", last);

                                    int lastUsedRow = last.Row;


                                    if (lastUsedRow != 1)
                                    {
                                        string tesster = (wsEXCEL.Cells[lastUsedRow, 1].Value2 + ".a").ToString();
                                        if (tesster.Length <= 2)
                                        {
                                            int l = lastUsedRow - 1;
                                        dataChecker:
                                            tesster = (wsEXCEL.Cells[l, 1].Value2 + ".a").ToString();
                                            if (tesster.Length <= 2)
                                            {
                                                l--;
                                                goto dataChecker;
                                            }
                                            else
                                            {
                                                lastUsedRow = l;
                                                tempLastusedrow = 0;
                                            }
                                        }
                                    }


                                    

                                    if (lastUsedRow != 1)
                                    {
                                        lastUsedRow = lastUsedRow + 2 - tempLastusedrow;
                                        tempLastusedrow = 0;
                                    }
                                    wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow, 1], wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + j] - Loc_Data[i]) - 1, a]].Value2 = wsCSV.Range[wsCSV.Cells[Loc_Data[i], 1], wsCSV.Cells[Loc_Data[i + j] - 1, a]].Value2;

                                        wsEXCEL.Cells[lastUsedRow, 1].EntireRow.Font.Bold = true;
                                    string cellValueCheck = (wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + j] - Loc_Data[i]) - 1, 2].Value2 + ".a").ToString();
                                    if (cellValueCheck == "Total.a")
                                    {
                                        wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + j] - Loc_Data[i]) - 1, 2].EntireRow.Font.Bold = true;
                                    }
                                    wsEXCEL.Columns.AutoFit();
                                    BorderAllSide(lastUsedRow, 1, lastUsedRow + (Loc_Data[i + j] - Loc_Data[i]) - 1, a);

                                    flag = false;

                                    int[] data = FindAllCustom("Program", lastUsedRow + 1, lastUsedRow + (Loc_Data[i + j] - Loc_Data[i]) - 1);
                                    for (int k = data.Length - 1; k >= 0; k--)
                                    {
                                        Range TempRange = (Range)wsEXCEL.Cells[data[k], 1];
                                        TempRange.EntireRow.Delete(Shift: XlDeleteShiftDirection.xlShiftUp);
                                        tempLastusedrow++;
                                       
                                    } 
                                }
                                else
                                {
                                    j++;
                                }
                            }
                            else
                            {
                                break;
                            }
                        } while (flag);
                        i = i + tempLastusedrow;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "MoveTablesMonthWise", "Exceptions Occured", "Yes", "395", ex.StackTrace, ex.ToString(), "Exception");
                return false;
                //return and error log record exception - LOG
            }

        }

        private string[] ExcelSheetNames()
        {
            LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "ExcelSheetNames", "Function started", null, "404", null, null, null);
            int count = 0;
            count = wbEXCEL.Sheets.Count;
            string[] Nam = new string[count];
            for (int i = 0; i < count; i++)
            {
                Worksheet oSheet = (Worksheet)wbEXCEL.Worksheets[i + 1];
                Nam[i] = oSheet.Name;
            }
            return Nam;
        }
        private void BorderAllSide(int RowFrom, int ColumnFrom, int RowTO, int ColumnTO)
        {

           // LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "BorderAllSide", "Function started", null, "417", null, null, null);
            var rowToBottomBorderizeRange = wsEXCEL.Range[wsEXCEL.Cells[RowFrom, ColumnFrom],
                wsEXCEL.Cells[RowTO, ColumnTO]];
            Borders border = rowToBottomBorderizeRange.Borders;
            border[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            border[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
        }
        private int[] FindAll(string key, string secondary, string ActualPDFfromSource)
        {
            try
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "FindAll", "Function started", null, "432", null, null, null);
                Range ur = wsCSV.UsedRange;
                Range r = wsCSV.Cells[1, ur.Columns.Count];
                r = r.get_End(XlDirection.xlToLeft);
                int lastCol = r.Column;

                string[] data = new string[0];
                int[] data1 = new int[0];
                int i = 0;
                Range currentFind = null;
                Range firstFind = null;
                object missing = System.Reflection.Missing.Value;
                Range Fruits;
                Fruits = excelCSV.get_Range("A1", "XFD1048576");
                currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, missing, missing);
                string loc1 = "";
                string loc = "";
                while (currentFind != null)
                {
                    // Keep track of the first range you find. 
                    if (firstFind == null)
                    {
                        firstFind = currentFind;
                        Array.Resize(ref data, data.Length + 1);
                        loc1 = currentFind.Address;
                        Array.Resize(ref data1, data1.Length + 1);
                        data1[i] = currentFind.Row;

                        loc1 = loc1.Replace("$", "");
                        data[i] = loc1;
                        i++;
                    }

                    // If you didn't move to a new range, you are done.
                    else if (currentFind.get_Address(XlReferenceStyle.xlA1)
                          == firstFind.get_Address(XlReferenceStyle.xlA1))
                    {
                        break;
                    }
                    currentFind = Fruits.FindNext(currentFind);


                    if (currentFind != null)
                    {
                        loc = currentFind.Address;

                        loc = loc.Replace("$", "");
                        currentFind.Activate();
                        if (firstFind != currentFind)
                        {
                            if (loc != loc1)
                            {
                                Array.Resize(ref data, data.Length + 1);

                                Array.Resize(ref data1, data1.Length + 1);
                                data1[i] = currentFind.Row;
                                data[i] = loc;
                                i++;
                            }
                        }
                    }
                }

                currentFind = null;
                firstFind = null;
                Fruits = excelCSV.get_Range("A" + data1[data1.Length - 1], "XFD1048576");
                key = "TAXES";

                currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, missing, missing);
                while (currentFind != null)
                {
                    // Keep track of the first range you find. 
                    if (firstFind == null)
                    {
                        firstFind = currentFind;
                        Array.Resize(ref data, data.Length + 1);
                        loc1 = currentFind.Address;
                        Array.Resize(ref data1, data1.Length + 1);
                        data1[i] = currentFind.Row;

                        loc1 = loc1.Replace("$", "");
                        data[i] = loc1;
                        i++;
                    }

                    // If you didn't move to a new range, you are done.
                    else if (currentFind.get_Address(XlReferenceStyle.xlA1)
                          == firstFind.get_Address(XlReferenceStyle.xlA1))
                    {
                        break;
                    }

                    currentFind = Fruits.FindNext(currentFind);

                    if (currentFind != null)
                    {
                        loc = currentFind.Address;

                        loc = loc.Replace("$", "");
                        currentFind.Activate();
                        if (firstFind != currentFind)
                        {
                            if (loc != loc1)
                            {
                                Array.Resize(ref data, data.Length + 1);

                                Array.Resize(ref data1, data1.Length + 1);
                                data1[i] = currentFind.Row;
                                data[i] = loc;
                                i++;
                            }
                        }
                    }
                }
                currentFind = null;
                firstFind = null;
                Fruits = excelCSV.get_Range(loc, "XFD1048576");
                key = "SGST";

                currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, missing, missing);
                while (currentFind != null)
                {
                    // Keep track of the first range you find. 
                    if (firstFind == null)
                    {
                        firstFind = currentFind;
                        Array.Resize(ref data, data.Length + 1);
                        loc1 = currentFind.Address;
                        Array.Resize(ref data1, data1.Length + 1);
                        data1[i] = currentFind.Row + 1;

                        loc1 = loc1.Replace("$", "");
                        data[i] = loc1;
                        i++;
                    }

                    // If you didn't move to a new range, you are done.
                    else if (currentFind.get_Address(XlReferenceStyle.xlA1)
                          == firstFind.get_Address(XlReferenceStyle.xlA1))
                    {
                        break;
                    }
                    currentFind = Fruits.FindNext(currentFind);

                    //string loc = "";
                    if (currentFind != null)
                    {
                        loc = currentFind.Address;

                        loc = loc.Replace("$", "");
                        currentFind.Activate();
                        if (firstFind != currentFind)
                        {
                            if (loc != loc1)
                            {
                                Array.Resize(ref data, data.Length + 1);

                                Array.Resize(ref data1, data1.Length + 1);
                                data1[i] = currentFind.Row + 1;
                                data[i] = loc;
                                i++;
                            }
                        }
                    }
                }
                return data1;

            }

            catch (Exception ex)
            {

                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "FindAll", "Exceptions Occured", "Yes", "603", ex.StackTrace, ex.ToString(), "Exception");
                File_Operations fs = new File_Operations();
                fs.ProcessedOrNot(ActualPDFfromSource, 2);
                return new int[0];
            }
        }

        private int[] FindAllCustom(string key, int from, int to)
        {
            try
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "FindAllCustom", "Function started", null, "612", null, null, null);

                Range ur = wsEXCEL.UsedRange;

                Range r = wsEXCEL.Cells[1, ur.Columns.Count];
                r = r.get_End(XlDirection.xlToLeft);
                int lastCol = r.Column;

                string[] data = new string[0];
                int[] data1 = new int[0];
                int i = 0;
                Range currentFind = null;
                Range firstFind = null;
                object missing = System.Reflection.Missing.Value;
                Range Fruits = (Range)wsEXCEL.Range[wsEXCEL.Cells[from, 1], wsEXCEL.Cells[to, 30]];

                currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, missing, missing);
                string loc1 = "";
                string loc = "";
                while (currentFind != null)
                {
                    // Keep track of the first range you find. 
                    if (firstFind == null)
                    {
                        firstFind = currentFind;
                        Array.Resize(ref data, data.Length + 1);
                        loc1 = currentFind.Address;
                        Array.Resize(ref data1, data1.Length + 1);
                        data1[i] = currentFind.Row;

                        loc1 = loc1.Replace("$", "");
                        data[i] = loc1;
                        i++;
                    }

                    // If you didn't move to a new range, you are done.
                    else if (currentFind.get_Address(XlReferenceStyle.xlA1)
                          == firstFind.get_Address(XlReferenceStyle.xlA1))
                    {
                        break;
                    }

                    currentFind = Fruits.FindNext(currentFind);


                    if (currentFind != null)
                    {
                        loc = currentFind.Address;

                        loc = loc.Replace("$", "");
                        currentFind.Activate();
                        if (firstFind != currentFind)
                        {
                            if (loc != loc1)
                            {
                                Array.Resize(ref data, data.Length + 1);

                                Array.Resize(ref data1, data1.Length + 1);
                                data1[i] = currentFind.Row;
                                data[i] = loc;
                                i++;
                            }
                        }
                    }
                }
                currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, missing, missing);
                return data1;
            }

            catch (Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "FindAllCustom", "Exceptions Occured", "Yes", "603", ex.StackTrace, ex.ToString(), "Exception");
                return new int[0];
            }
}
        private int GetLastColumn(int row)
        {

           // LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "GetLastColumn", "Function started", null, "690", null, null, null);

            int i1 = 1;

            while (row > 0)
            {
                string temp = wsCSV.Cells[row, i1].Value + ".aaa";
                if (String.IsNullOrEmpty(temp) || String.IsNullOrWhiteSpace(temp) || temp.Equals(".aaa") == true)
                {
                    break;
                }
                i1++;
            }
            return i1 - 1;
        }

        #endregion

        #region Phase 2 - Data Processing
        //Phase 2: Over View.

        private string[,] DataController(string PDFFile)
        {
            try
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "DataController", "Function started. PDFFile : " + PDFFile, null, "713", null, null, null);

                string[,] PDFExtracedData = DataExtracter(PDFFile);
                if (PDFExtracedData.GetLength(0) > 0)
                {
                    return PDFExtracedData;
                }
                else
                {
                    return new string[0, 0];
                }
            }

            catch (Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "FindAllCustom", "Exceptions Occured. PDFFile : "+ PDFFile, "Yes", "603", ex.StackTrace, ex.ToString(), "Exception");
                return new string[0,0];
            }
        }
        private string[,] DataExtracter(string filename)
        {
            try
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "DataExtracter", "Function started. filename : "  + filename, null, "736", null, null, null);

                StringBuilder text1 = new StringBuilder();
                if (File.Exists(filename))
                {
                    using (PdfReader reader = new PdfReader(filename))
                    {
                        for (int i = 1; i <= reader.NumberOfPages; i++)
                        {
                            text1.Append("::::::::::");
                            text1.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                        }
                    }
                }
                string[,] review = OverView(text1.ToString());
                return review;
            }

            catch (Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "DataExtracter", "Exceptions Occured", "Yes", "756", ex.StackTrace, ex.ToString(), "Exception");
                return new string[0, 0]; 
            }

        }
        private string[,] OverView(string ActualData)
        {
           // LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "OverView", "Function started. ActualData : "+ ActualData, null, "762", null, null, null);
            LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "OverView", "Function started. ", null, "762", null, null, null);
            string[] _Month = new string[0];
            string[] _Year = new string[0];
            // filter remove table data          
            wsEXCEL = wbEXCEL.Worksheets["Sheet1"];
            wsEXCEL.Name = "Overview";
            wsEXCEL.Activate();
            string delimeter1 = "::::::::::";
            string delimeter2 = "\n";
            string[] PageData = ActualData.Split(new string[] { delimeter1 }, StringSplitOptions.None);
            //string[] FinalData = new string[0];
            int i = 0;
            int j = 1;
            int sheetnum = 0;
            int count = 0;
            count = wbEXCEL.Sheets.Count;
            string[] sheetName = new string[count];
            for (int ii = 0; ii < count; ii++)
            {
                Worksheet oSheet = (Worksheet)wbEXCEL.Worksheets[ii + 1];
                sheetName[ii] = oSheet.Name;
            }
            for (int EachPageData = 1; EachPageData < PageData.Length; EachPageData++)
            {
                wsEXCEL.Cells[j, 1].EntireRow.Font.Bold = true;
                string[] RawData = PageData[EachPageData].Split(new string[] { delimeter2 }, StringSplitOptions.None);
                bool DataFound = true;
                for (int eachLine = 0; eachLine < RawData.Length; eachLine++)
                {
                LabelReviertBack:

                    if (RawData[eachLine].Contains("Channel Program Program Title"))
                    {
                        while (((RawData[eachLine].Contains("JAN") || RawData[eachLine].Contains("FEB") || RawData[eachLine].Contains("MAR") ||
                              RawData[eachLine].Contains("APR") || RawData[eachLine].Contains("MAY") || RawData[eachLine].Contains("JUN") ||
                              RawData[eachLine].Contains("JUL") || RawData[eachLine].Contains("AUG") || RawData[eachLine].Contains("SEP") ||
                              RawData[eachLine].Contains("OCT") || RawData[eachLine].Contains("NOV") || RawData[eachLine].Contains("DEC")) &&
                              RawData[eachLine].Contains("Total")) == false)
                        {
                            if (RawData[eachLine].Contains("Page") && RawData[eachLine].Contains("of"))
                            {
                                break;
                            }
                            else
                            {
                                if (eachLine < RawData.Length - 1)
                                {
                                    eachLine++;
                                }
                                else
                                {
                                    break;
                                }

                            }

                        }
                        DataFound = false;
                        goto LabelReviertBack;
                    }
                    else if (RawData[eachLine].Contains("TAXES Rate Cost"))
                    {
                        while (RawData[eachLine].Contains("SGST") == false)
                        {
                            if (RawData[eachLine].Contains("Page") && RawData[eachLine].Contains("of"))
                            {
                                break;
                            }
                            else
                            {
                                if (eachLine < RawData.Length - 1)
                                {
                                    eachLine++;
                                }
                                else
                                {
                                    break;
                                }

                            }
                        }
                        if (eachLine < RawData.Length - 1)
                        {
                            eachLine++;
                        }
                        DataFound = false;
                        goto LabelReviertBack;
                    }
                    else if ((RawData[eachLine].Contains("JAN") || RawData[eachLine].Contains("FEB") || RawData[eachLine].Contains("MAR") ||
                              RawData[eachLine].Contains("APR") || RawData[eachLine].Contains("MAY") || RawData[eachLine].Contains("JUN") ||
                              RawData[eachLine].Contains("JUL") || RawData[eachLine].Contains("AUG") || RawData[eachLine].Contains("SEP") ||
                              RawData[eachLine].Contains("OCT") || RawData[eachLine].Contains("NOV") || RawData[eachLine].Contains("DEC")) &&
                              RawData[eachLine].Contains("Total"))
                    {
                        eachLine++;
                        DataFound = true;
                    }
                    else if (RawData[eachLine].Contains("Page") && RawData[eachLine].Contains("of"))
                    {
                        DataFound = true;
                    }
                    else
                    {
                        DataFound = true;
                    }
                    if (DataFound == true)
                    {
                        if (eachLine < RawData.Length - 1)
                        { 
                            wsEXCEL.Cells[j, 1].Value2 = RawData[eachLine].Trim();
                            if (RawData[eachLine].Contains("Channel:"))
                            {
                                int pos = RawData[eachLine].IndexOf("month of", StringComparison.CurrentCultureIgnoreCase);
                                string a = RawData[eachLine].Substring(pos + 1 + "month of".Length, 3);
                                string b = RawData[eachLine].Substring(pos + 1 + "month of".Length+ 4,4);
                                if(!Array.Exists(_Month, element => element == a))
                                {
                                    Array.Resize(ref _Month, _Month.Length + 1);
                                    Array.Resize(ref _Year, _Year.Length + 1);
                                    _Month[_Month.Length - 1] = a;
                                    _Year[_Year.Length - 1] = b;
                                }
                                //j = j + 1;
                                //string Sht2_B1 = "#'" + a + "'!A1";
                                //wsEXCEL.Cells[j, 1].Formula = "=HYPERLINK(\"" + Sht2_B1 + "\", \"" + a + "\")"; ;
                                sheetnum++;
                                j = j + 1;
                                wsEXCEL.Cells[j, 1].Value2 = "";
                            }
                            else if (eachLine + 1 < RawData.Length - 1)
                            {
                                if (RawData[eachLine + 1].Contains("Total amount payable: "))
                                {
                                    if (sheetnum < sheetName.Length)
                                    {
                                        //j = j + 1;
                                        //string Sht2_B1 = "#'" + sheetName[sheetnum] + "'!A1";
                                        //wsEXCEL.Cells[j, 1].Formula = "=HYPERLINK(\"" + Sht2_B1 + "\", \"" + sheetName[sheetnum] + "\")"; ;
                                        sheetnum++;
                                        j = j + 1;
                                        wsEXCEL.Cells[j, 1].Value2 = "";
                                    }
                                }
                            }
                            i++;
                            j++;
                        }
                    }
                }
                j++;

                wsEXCEL.Cells[j, 1].Value2 = "";
            }
            string[,] _Month_Year = new string[_Month.Length, _Year.Length];
            try
            {
                count = 0;
                count = wbEXCEL.Sheets.Count;
                wbEXCEL.Worksheets["Overview"].Move(After: wbEXCEL.Worksheets[count]);
                wsEXCEL.SaveAs(path);
                
                if(_Month.Length == _Year.Length)
                {
                    try
                    {
                        _Month_Year = TwoDArrayFromTwoOneD(_Month, _Year);
                    }
                    catch(Exception ex)
                    {
                        LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "OverView", "Exceptions Occured", "Yes", "756", ex.StackTrace, ex.ToString(), "Exception");

                        return _Month_Year;
                    }
                    
                }
                 
            }   
            catch(Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "OverView", "Exceptions Occured", "Yes", "756", ex.StackTrace, ex.ToString(), "Exception");

                //Console.WriteLine("CoreFunction 2: 789" + ex + "\n" + ex.StackTrace); ;
            }
            return _Month_Year;

        }

        private static string[,] TwoDArrayFromTwoOneD(string[] Mat1, string[] Mat2)
        {
            LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "TwoDArrayFromTwoOneD", "Function started", null, "950", null, null, null);
            string[,] newMat = new string[2, Mat2.Length]; 
            for (var j = 0; j < Mat1.Length; j++)
            {
                newMat[0, j] = Mat1[j];
                newMat[1, j] = Mat2[j];
            } 
            return newMat;
        }
        #endregion

        #region Phase 3 - Memory Management
        private void ReleaseMemory()
        {
            try
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "ReleaseMemory", "Function started", null, "966", null, null, null);

                wbCSV.Close(true);
                wbEXCEL.Close(true);
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //if (wbCSV != null)
                //    wbCSV.Close(true, Missing.Value, Missing.Value);

                //if (wsCSV != null)
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(wsCSV);

                //wsCSV = null;

                //if (wbCSV != null)
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(wbCSV);
                //wbCSV = null;



                //if (wbEXCEL  != null)
                //    wbEXCEL.Close(true, Missing.Value, Missing.Value);

                //if (wsEXCEL != null)
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(wsEXCEL);

                //wsEXCEL = null;

                //if (wbEXCEL != null)
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(wbEXCEL);
                //wbEXCEL = null;


            }
            catch(Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "OverView", "Exceptions Occured", "Yes", "1003", ex.StackTrace, ex.ToString(), "Exception");
 
            }



        }
        #endregion 

        #region phase 0.1 - CSV Table Data correction

       
        void Data_MON_FIND_REPLACE(int row)
        {
            try
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "Data_MON_FIND_REPLACE", "Function started. Row : " + row, null, "1021", null, null, null);

                string CelltoStart = (wsCSV.Cells[row, 11].Value2 + ".a").ToString().ToUpper();
                if (CelltoStart.Contains("NET COST"))
                {
                    string MOnFirst = (wsCSV.Cells[row, 13].Value2 + ".a").ToString().ToUpper();
                    if (MOnFirst.Contains("TUE"))
                    {
                        string data = ("a"+wsCSV.Cells[row, 12].Value2).ToString().ToUpper().Substring(wsCSV.Cells[row, 12].Value2.ToString().ToUpper().IndexOfAny("0123456789".ToCharArray()));
                        wsCSV.Cells[row, 12].Value2 = "MON" + data;

                    }
                    for (int i = 11; i < 44; i++)
                    {

                        string findString = (wsCSV.Cells[row, i].Value2 + ".a").ToString().ToUpper();
                        if (findString.Contains("SUN"))
                        {
                            string data = ("a" + wsCSV.Cells[row, i + 1].Value2).ToString().ToUpper().Substring(wsCSV.Cells[row, i + 1].Value2.ToString().ToUpper().IndexOfAny("0123456789".ToCharArray()));
                            wsCSV.Cells[row, i + 1].Value2 = "MON" + data;
                        }
                    }
                }
            }
            catch(Exception ex)
            { 
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "Data_MON_FIND_REPLACE", "Exceptions Occured", "Yes", "1043", ex.StackTrace, ex.ToString(), "Exception");

            }
        }
        void Data_Correction_MON(int[] loc)
        {
            try
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "Data_Correction_MON", "Function started", null, "1055", null, null, null);

                for (int i = 0; i < loc.Length; i++)
                {
                    //string temp1 = (wsCSV.Cells[loc[i], 11].Value2 + ".a").ToString().ToUpper();
                    //Data_MON_FIND_REPLACE(loc[i]);
                    string CelltoStart = (wsCSV.Cells[loc[i], 11].Value2 + ".a").ToString().ToUpper();
                    if (CelltoStart.Contains("NET COST"))
                    {
                        int[] j = FindAllCustomforMON("sun", loc[i]);
                        for (int k = 0; k < j.Length; k++)
                        {
                            string data = ("a" + wsCSV.Cells[loc[i], j[k] + 1].Value2).ToString().ToUpper().Substring(("a"+wsCSV.Cells[loc[i], j[k] + 1].Value2).ToString().ToUpper().IndexOfAny("0123456789".ToCharArray()));
                            wsCSV.Cells[loc[i], j[k] + 1].Value2 = "MON" + data;
                            
                        }

                        string MOnFirst = (wsCSV.Cells[loc[i], 13].Value2 + ".a").ToString().ToUpper();
                        if (MOnFirst.Contains("TUE"))
                        {
                            string data = ("a" + wsCSV.Cells[loc[i], 12].Value2).ToString().ToUpper().Substring(("a" + wsCSV.Cells[loc[i], 12].Value2).ToString().ToUpper().IndexOfAny("0123456789".ToCharArray()));
                            wsCSV.Cells[loc[i], 12].Value2 = "MON" + data; 
                        }
                    } 
                }
            }
            catch (Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "Data_Correction_MON", "Exceptions Occured", "Yes", "1065", ex.StackTrace, ex.ToString(), "Exception");
            }
        }


        private int[] FindAllCustomforMON(string key, int from)
        {
            try
            {
                Range ur = wsCSV.UsedRange;

                Range r = wsCSV.Cells[1, ur.Columns.Count];
                r = r.get_End(XlDirection.xlToLeft);
                int lastCol = r.Column;

                string[] data = new string[0];
                int[] data1 = new int[0];
                int i = 0;
                Range currentFind = null;
                Range firstFind = null;
                object missing = System.Reflection.Missing.Value;
                Range Fruits = (Range)wsCSV.Range[wsCSV.Cells[from, 11], wsCSV.Cells[from, 45]];

                currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, missing, missing);
                string loc1 = "";
                string loc = "";
                while (currentFind != null)
                {
                    // Keep track of the first range you find. 
                    if (firstFind == null)
                    {
                        firstFind = currentFind;
                        Array.Resize(ref data, data.Length + 1);
                        loc1 = currentFind.Address;
                        Array.Resize(ref data1, data1.Length + 1);
                        data1[i] = currentFind.Column;

                        loc1 = loc1.Replace("$", "");
                        data[i] = loc1;
                        i++;
                    }

                    // If you didn't move to a new range, you are done.
                    else if (currentFind.get_Address(XlReferenceStyle.xlA1)
                          == firstFind.get_Address(XlReferenceStyle.xlA1))
                    {
                        break;
                    }

                    currentFind = Fruits.FindNext(currentFind);


                    if (currentFind != null)
                    {
                        loc = currentFind.Address;

                        loc = loc.Replace("$", "");
                        currentFind.Activate();
                        if (firstFind != currentFind)
                        {
                            if (loc != loc1)
                            {
                                Array.Resize(ref data, data.Length + 1);

                                Array.Resize(ref data1, data1.Length + 1);
                                data1[i] = currentFind.Column;
                                data[i] = loc;
                                i++;
                            }
                        }
                    }
                }
                currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, missing, missing);
                return data1;
            }

            catch (Exception ex)
            {
                LogHandeler.ModuleLogFile("---", "---", "---", "Core_Operation2", "FindAllCustom", "Exceptions Occured", "Yes", "603", ex.StackTrace, ex.ToString(), "Exception");
                return new int[0];
            }
        }
        #endregion
    }
}
