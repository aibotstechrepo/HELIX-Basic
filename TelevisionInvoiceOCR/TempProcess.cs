using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Text; 
using _Excel = Microsoft.Office.Interop.Excel;

namespace TelevisionInvoiceOCR
{
    class TempProcess      {

        #region operaion Main
        
        
        //for CSV
        public _Application excelCSV = new _Excel.Application();
        public Workbook wbCSV;  
        public Worksheet wsCSV;

        //for EXCEL
        public _Application excelEXCEL = new _Excel.Application();
        public Workbook wbEXCEL;
        public Worksheet wsEXCEL;

        string path = "d:\\1.xlsx";

        int tempLastusedrow = 0;

        #endregion
        internal void Extract(string fileName, string OutfileName)
        {
            //Form1.progressBar1.Value = 10;
            //Form1.label6.Text = "10";
            NewExcel();
            #region Phase1

            path = OutfileName;
            fileName = "D:\\1.csv";
            //Form1.progressBar1.Value = 20;
            //Form1.label6.Text = "20";
            Open_Excel(fileName);
            //Form1.progressBar1.Value = 30;
            //Form1.label6.Text = "30";
            int[] loc = FindAll("Program", "Channel");
            //Form1.progressBar1.Value = 50;
            //Form1.label6.Text = "50";
            Transfer(loc);
            //Form1.progressBar1.Value = 70;
            //Form1.label6.Text = "70";
            #endregion


            #region Phase2 
            DataController("D:\\1.pdf");
            //Form1.progressBar1.Value = 90;
            //Form1.label6.Text = "90";
            //DataController("D:\\Final_est_pages.pdf");
            #endregion

            #region Phase3
            //Form1.progressBar1.Value = 100;
            //Form1.label6.Text = "100";
            ReleaseMemory();
            #endregion

        }


        #region Phase1



        //Create new Excelfile
        private void NewExcel()
        {
            wbEXCEL = excelEXCEL.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            wsEXCEL = wbEXCEL.Worksheets[1];
            excelEXCEL.Visible = true;
        }

        //Open CSVFile
        private void Open_Excel(string Path)
        {
            wbCSV = excelCSV.Workbooks.Open(Path);
            wsCSV = wbCSV.Worksheets[1];
            //wsCSV.Activate();
            excelCSV.Visible = false;
        }

        //Find CSVFile Keyword 
       

        //actual operation
        private void Transfer(int[] Loc_Data)
        { 
            //for(int i=0; i<1;i++)
            for (int i = 0; i < Loc_Data.Length-1; i++)
            {


                //int last = 0;
                //Range ur = wsCSV.UsedRange;
                //string b1 = ur.Address;
                //Range r = wsCSV.Cells[i+1, ur.Columns.Count];
                //r = r.get_End(XlDirection.xlToLeft); 
                //int lastCol = r.Column;
                //string b = r.Address;
                //if (wsCSV.Cells[Loc_Data[i] + 1, 1 == "Total");

                if (i == Loc_Data.Length - 2)
                {
                    string temp1 = wsCSV.Cells[Loc_Data[i],1].Value2.ToString().ToUpper();
                    if ((temp1 == "TAXES"))
                    {
                        wbEXCEL.Sheets.Add(After: wbEXCEL.Sheets[wbEXCEL.Sheets.Count]);
                        wsEXCEL = wbEXCEL.Worksheets[wbEXCEL.Sheets.Count];
                        wsEXCEL.Name = "TAXES";
                        wsEXCEL.Activate();

                        Range last = wsEXCEL.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                        Range range = wsEXCEL.get_Range("A1", last);

                        int lastUsedRow = last.Row;
                        int a = GetLastColumn(Loc_Data[i]); 

                        if (lastUsedRow != 1)
                        {
                            lastUsedRow = lastUsedRow + 2 - tempLastusedrow;
                            tempLastusedrow = 0;
                        }
                         
                        wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow, 1], wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + 1] - Loc_Data[i]) - 1, a]].Value2 = wsCSV.Range[wsCSV.Cells[Loc_Data[i], 1], wsCSV.Cells[Loc_Data[i + 1] - 1, a]].Value2;
                        wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow+1, 2], wsEXCEL.Cells[lastUsedRow + 1, 2]].NumberFormat = "###,##%";
                        wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow+2, 2], wsEXCEL.Cells[lastUsedRow + 2, 2]].NumberFormat = "###,##%";

                        wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow + 1, 4], wsEXCEL.Cells[lastUsedRow + 1, 4]].NumberFormat = "###,##%";
                        wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow + 2, 4], wsEXCEL.Cells[lastUsedRow + 1, 4]].NumberFormat = "###,##%";

                        //last = Loc_Data[i + 1] - Loc_Data[i];
                        wsEXCEL.Cells[lastUsedRow, 1].EntireRow.Font.Bold = true;
                         
                        wsEXCEL.Columns.AutoFit();
                        BorderAllSide(lastUsedRow, 1, lastUsedRow + (Loc_Data[i + 1] - Loc_Data[i]) - 1, a);


                    }
                }
                

                string temp = wsCSV.Cells[Loc_Data[i+1]-1, 2].Value2.ToString().ToUpper();
                if ((temp == "TOTAL"))
                {
                    //System.Windows.Forms.MessageBox.Show("Month = "+ wsCSV.Cells[Loc_Data[i + 1] - 1, 1].Value2.ToString().ToUpper() + " \t Row  = "+i);
                    string[] SheetAvaialble = ExcelSheetNames();

                    if (Array.Exists(SheetAvaialble, element => element == wsCSV.Cells[Loc_Data[i + 1] - 1, 1].Value2.ToString().ToUpper()))
                    {
                        wsEXCEL = wbEXCEL.Worksheets[wsCSV.Cells[Loc_Data[i + 1] - 1, 1].Value2.ToString().ToUpper()];

                        wsEXCEL.Activate();
                    }
                    else
                    {
                        wbEXCEL.Sheets.Add(After: wbEXCEL.Sheets[wbEXCEL.Sheets.Count]);
                        wsEXCEL = wbEXCEL.Worksheets[wbEXCEL.Sheets.Count];
                        wsEXCEL.Name = wsCSV.Cells[Loc_Data[i + 1] - 1, 1].Value2.ToString().ToUpper();

                        wsEXCEL.Activate();
                    }

                    Range last = wsEXCEL.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                    Range range = wsEXCEL.get_Range("A1", last);

                    int lastUsedRow = last.Row;

                    //int a = last.Column;

                    int a = GetLastColumn(Loc_Data[i]);

                    //Range rangeFrom = (Range)wsCSV.Range[wsCSV.Cells[Loc_Data[i], 1], wsCSV.Cells[Loc_Data[i + 1] - 1, a]];
                    //string a1 = rangeFrom.Address;

                    //Range rangeTo = (Range)wsEXCEL.Range[wsEXCEL.Cells[1, 1], wsEXCEL.Cells[Loc_Data[i + 1] - Loc_Data[i], a]];
                    //string b1 = rangeTo.Address;

                    //wsEXCEL = wbEXCEL.Worksheets[wbEXCEL.Sheets.Count];
                    //wsEXCEL.Activate();
                    if (lastUsedRow != 1)
                    {
                        lastUsedRow = lastUsedRow + 2 - tempLastusedrow;
                        tempLastusedrow = 0;
                    }



                    wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow, 1], wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + 1] - Loc_Data[i])-1, a]].Value2 = wsCSV.Range[wsCSV.Cells[Loc_Data[i], 1], wsCSV.Cells[Loc_Data[i + 1] - 1, a]].Value2;
                    //last = Loc_Data[i + 1] - Loc_Data[i];
                    wsEXCEL.Cells[lastUsedRow, 1].EntireRow.Font.Bold = true;
                    string cellValueCheck = (wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + 1] - Loc_Data[i]) - 1, 2].Value2 + ".a").ToString();
                    if (cellValueCheck == "Total.a")
                    {
                        wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + 1] - Loc_Data[i]) - 1, 2].EntireRow.Font.Bold = true;
                    }


                    wsEXCEL.Columns.AutoFit();
                    BorderAllSide(lastUsedRow, 1, lastUsedRow + (Loc_Data[i + 1] - Loc_Data[i])-1,a);
                    //Range from = wsCSV.Range["A1:A100"];


                    //rangeFrom.Copy(rangeTo);
                    //from.Copy(to);
                }
                else
                {
                    bool flag = true;
                    int j = 2;
                    do
                    {
                        if ((i + j) < Loc_Data.Length - 1)
                        {
                            temp = wsCSV.Cells[Loc_Data[i + j] - 1, 2].Value2.ToString().ToUpper();
                            if ((temp == "TOTAL"))
                            {

                                string[] SheetAvaialble = ExcelSheetNames();

                                if (Array.Exists(SheetAvaialble, element => element == wsCSV.Cells[Loc_Data[i + j] - 1, 1].Value2.ToString().ToUpper()))
                                {
                                    wsEXCEL = wbEXCEL.Worksheets[wsCSV.Cells[Loc_Data[i + j] - 1, 1].Value2.ToString().ToUpper()];

                                    wsEXCEL.Activate();
                                }
                                else
                                {
                                    wbEXCEL.Sheets.Add(After: wbEXCEL.Sheets[wbEXCEL.Sheets.Count]);
                                    wsEXCEL = wbEXCEL.Worksheets[wbEXCEL.Sheets.Count];
                                    wsEXCEL.Name = wsCSV.Cells[Loc_Data[i + j] - 1, 1].Value2.ToString().ToUpper();

                                    wsEXCEL.Activate();
                                }

                                Range last = wsEXCEL.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                                Range range = wsEXCEL.get_Range("A1", last);

                                int lastUsedRow = last.Row;

                                //int a = last.Column;

                                int a = GetLastColumn(Loc_Data[i]);

                                //Range rangeFrom = (Range)wsCSV.Range[wsCSV.Cells[Loc_Data[i], 1], wsCSV.Cells[Loc_Data[i + 1] - 1, a]];
                                //string a1 = rangeFrom.Address;

                                //Range rangeTo = (Range)wsEXCEL.Range[wsEXCEL.Cells[1, 1], wsEXCEL.Cells[Loc_Data[i + 1] - Loc_Data[i], a]];
                                //string b1 = rangeTo.Address;

                                //wsEXCEL = wbEXCEL.Worksheets[wbEXCEL.Sheets.Count];
                                //wsEXCEL.Activate();
                                if (lastUsedRow != 1)
                                {
                                    lastUsedRow = lastUsedRow + 2 - tempLastusedrow;
                                    tempLastusedrow = 0;
                                }



                                wsEXCEL.Range[wsEXCEL.Cells[lastUsedRow, 1], wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + j] - Loc_Data[i]) - 1, a]].Value2 = wsCSV.Range[wsCSV.Cells[Loc_Data[i], 1], wsCSV.Cells[Loc_Data[i + j] - 1, a]].Value2;
                                //last = Loc_Data[i + 1] - Loc_Data[i];
                                wsEXCEL.Cells[lastUsedRow, 1].EntireRow.Font.Bold = true;
                                string cellValueCheck = (wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + j] - Loc_Data[i]) - 1, 2].Value2 + ".a").ToString();
                                if (cellValueCheck == "Total.a")
                                {
                                    wsEXCEL.Cells[lastUsedRow + (Loc_Data[i + j] - Loc_Data[i]) - 1, 2].EntireRow.Font.Bold = true;
                                }


                                wsEXCEL.Columns.AutoFit();
                                BorderAllSide(lastUsedRow, 1, lastUsedRow + (Loc_Data[i + j] - Loc_Data[i]) - 1, a);
                                //Range from = wsCSV.Range["A1:A100"];

                                 
                                //rangeFrom.Copy(rangeTo);
                                //from.Copy(to);
                                flag = false;

                                int[] data = FindAllCustom("Program", lastUsedRow + 1, lastUsedRow + (Loc_Data[i + j] - Loc_Data[i]) - 1);
                                for(int k=data.Length-1; k>=0; k--)                                
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
                    
                }
            }
            
        }

        private string[] ExcelSheetNames()
        {
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
        private int[] FindAll(string key, string secondary)
        {
            //last used row and column number
            //int lastRow = wsCSV.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            //int lastCol = wsCSV.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;

            //Range from and to entire excel
            //Range last = wsCSV.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            //Range range = wsCSV.get_Range("A1", last);
            //string a = range.Address;

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

            // currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, missing, missing);
            //  currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, true, missing, missing);
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

                //currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                //currentFind.Font.Bold = true;

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



                //System.Windows.MessageBox.Show(loc);
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

                //currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                //currentFind.Font.Bold = true;

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
                            data1[i] = currentFind.Row;
                            data[i] = loc;
                            i++;
                        }
                    }
                }



                //System.Windows.MessageBox.Show(loc);
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

                //currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                //currentFind.Font.Bold = true;

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



                //System.Windows.MessageBox.Show(loc);
            }





            return data1;

        }

        private int[] FindAllCustom(string key, int from, int to)
        {
            //last used row and column number
            //int lastRow = wsCSV.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            //int lastCol = wsCSV.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;

            //Range from and to entire excel
            //Range last = wsCSV.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            //Range range = wsCSV.get_Range("A1", last);
            //string a = range.Address;

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
            //Fruits = wsEXCEL.get_Range(wsEXCEL.Cells[from, 1], wsEXCEL.Cells[to, 30]); 

            currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, missing, missing);

            // currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, missing, missing);
            //  currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, true, missing, missing);
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

                //currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                //currentFind.Font.Bold = true;

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



                //System.Windows.MessageBox.Show(loc);
            }

            //currentFind = null;
            //firstFind = null;
            //Fruits = excelCSV.get_Range("A" + data1[data1.Length - 1], "XFD1048576");
            //key = "TAXES";

            currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, missing, missing);
            //while (currentFind != null)
            while (1!= 1)
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

                //currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                //currentFind.Font.Bold = true;

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
                            data1[i] = currentFind.Row;
                            data[i] = loc;
                            i++;
                        }
                    }
                }



                //System.Windows.MessageBox.Show(loc);
            }
            //currentFind = null;
            //firstFind = null;
            //Fruits = excelCSV.get_Range(loc, "XFD1048576");
            //key = "SGST";

            currentFind = Fruits.Find(key, missing, XlFindLookIn.xlValues, XlLookAt.xlWhole, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, missing, missing);
            //while (currentFind != null)
            while (1!= 1)
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

                //currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                //currentFind.Font.Bold = true;

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
                //System.Windows.MessageBox.Show(loc);
            } 
            return data1;

        }
        private int GetLastColumn(int row)
        {

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
            return i1-1;
        }

        #endregion

        #region Phase 2
        //Phase 2: Over View.

        private void DataController(string PDFFile)
        {
            string PDFExtracedData = DataExtracter(PDFFile);
        }
        private string DataExtracter(string filename)
        {
            //try
            //{ 
            StringBuilder text1 = new StringBuilder();
            //using (PdfReader reader = new PdfReader(@".pdf"))
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
            // System.Windows.Forms.MessageBox.Show(text1.ToString());

            OverView(text1.ToString());
            return text1.ToString();

        }
        private string[] OverView(string ActualData)
        {
            // filter remove table data
            //wsEXCEL = wbEXCEL.Worksheets[wbEXCEL.Sheets.Count];
            wsEXCEL = wbEXCEL.Worksheets["Sheet1"];
            wsEXCEL.Name = "Overview";
            wsEXCEL.Activate();
            //string fil = "";
            string delimeter1 = "::::::::::";
            string delimeter2 = "\n";
            string[] PageData = ActualData.Split(new string[] { delimeter1 }, StringSplitOptions.None);
            string[] FinalData = new string[0];
            int i = 0;
            int j = 1;
            int sheetnum =0;
            int count = 0;
            count = wbEXCEL.Sheets.Count; 
            string[] sheetName = new string[count];
            for (int ii = 0; ii < count; ii++)
            {
                Worksheet oSheet = (Worksheet)wbEXCEL.Worksheets[ii + 1];
                sheetName[ii] = oSheet.Name;
            }
            for (int EachPageData = 1; EachPageData< PageData.Length; EachPageData++)
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
                                if(eachLine < RawData.Length-1)
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
                    else if (RawData[eachLine].Contains("TAXES Rate Cost Rate"))
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
                            Array.Resize(ref FinalData, FinalData.Length + 1);
                            FinalData[i] = RawData[eachLine].Trim();
                            wsEXCEL.Cells[j, 1].Value2 = RawData[eachLine].Trim(); 
                            if(RawData[eachLine].Contains("Channel:"))
                            {
                                int pos = RawData[eachLine].IndexOf("month of", StringComparison.CurrentCultureIgnoreCase);
                                string a = RawData[eachLine].Substring(pos + 1 + "month of".Length, 3);


                                //if (sheetnum < sheetName.Length)
                                //{
                                    j = j + 1;
                                    string Sht2_B1 = "#'" + a + "'!A1";
                                    wsEXCEL.Cells[j, 1].Formula = "=HYPERLINK(\"" + Sht2_B1 + "\", \"" + a + "\")"; ;
                                    sheetnum++;
                                    j = j + 1;
                                    wsEXCEL.Cells[j, 1].Value2 = "";
                                //}
                            }
                            else if (eachLine+1 < RawData.Length-1)
                            {
                                if (RawData[eachLine + 1].Contains("Total amount payable: "))
                                {
                                    if (sheetnum < sheetName.Length)
                                    {
                                        j = j + 1;
                                        string Sht2_B1 = "#'" + sheetName[sheetnum] + "'!A1";
                                        wsEXCEL.Cells[j, 1].Formula = "=HYPERLINK(\"" + Sht2_B1 + "\", \"" + sheetName[sheetnum] + "\")"; ;
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
                //string Sht2_B1 = "#'"+ sheetName[sheetnum] + "'!A1";
                //wsEXCEL.Cells[j, 1].Formula = "=HYPERLINK(\"" + Sht2_B1 + "\", \"" + sheetName[sheetnum] + "\")"; ;
                //sheetnum++;
                j ++;
               
                wsEXCEL.Cells[j, 1] .Value2 = "";

            }
            try
            {
                count = 0;
                count = wbEXCEL.Sheets.Count;
                wbEXCEL.Worksheets["Overview"].Move(After: wbEXCEL.Worksheets[count]); 
                wsEXCEL.SaveAs(path);
            }
            catch
            {
                ;
            }
            //wsEXCEL.SaveAs(path);
            return FinalData; 
        }


        #endregion

        #region Phase3
         private void ReleaseMemory()
        {
            GC.Collect(); 
            GC.WaitForPendingFinalizers();
        }
        #endregion

    }
}
