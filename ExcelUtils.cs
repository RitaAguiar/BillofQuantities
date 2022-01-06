#region namespaces

using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using Autodesk.Revit.UI;

#endregion //namespaces

namespace BillofQuantities
{
    public class ExcelUtils
    {
        public static Excel.Application excel = null;
        public static Excel.Workbook workbook = null;

        public static void LauchExcel()
        {
            #region Launch or access Excel via COM Interop and create worksheet

            excel = new Excel.Application(); // access excel
            excel.Visible = true; // excel document visible/open
            workbook = excel.Workbooks.Add(Missing.Value); // access new workbook .xls

            #endregion Launch or access Excel via COM Interop and create worksheet
        }

        public static void PreventInteraction()
        {
            excel.Interactive = false; //prevents user from interacting with Excel
        }

        public static void CreateInstancesSpreadsheet(List<string> paramNamesEI, List<EI> EISSorted)
        {
            #region Excel Sheet "Instances"

            Excel.Worksheet worksheetEI = workbook.Sheets.get_Item(1) as Excel.Worksheet; // gets first worksheet
            worksheetEI.Name = "Instances"; // worksheetEI name
            worksheetEI.EnableSelection = Excel.XlEnableSelection.xlNoSelection;

            #endregion Excel Sheet "Instances"

            #region HEADER - 1ST ROW - Instance Elements

            int columnEI = 1;
            int rowEI = 1;

            worksheetEI.Cells[rowEI, columnEI++] = "ID";
            worksheetEI.Cells[rowEI, columnEI++] = "IsType";
            worksheetEI.Cells[rowEI, columnEI++] = "Category";
            worksheetEI.Cells[rowEI, columnEI++] = "Type Name";
            worksheetEI.Cells[rowEI, columnEI++] = "Type Name Id";
            worksheetEI.Cells[rowEI, columnEI++] = "Family Name"; // column 6

            foreach (string paramName in paramNamesEI) // writes paramName after column 7
            {
                worksheetEI.Cells[1, columnEI] = paramName; 
                ++columnEI;
            }

            // Fitting
            var range_rowEI = worksheetEI.get_Range("A1", "A1").EntireRow; // all cels of the first row
            range_rowEI.Font.Bold = true;
            range_rowEI.EntireColumn.AutoFit();

            #endregion HEADER - 1ST ROW - Instance Elements

            #region OTHER ROWS - Instance Elements

            // Filling
            foreach (var ListEI in EISSorted)
            {
                ++rowEI; // +1 row
                columnEI = 1;

                // writes each element of ListEI in a cell on each column
                worksheetEI.Cells[rowEI, columnEI++] = ListEI.ID;
                worksheetEI.Cells[rowEI, columnEI++] = ListEI.IsType;
                worksheetEI.Cells[rowEI, columnEI++] = ListEI.CategoryName;
                worksheetEI.Cells[rowEI, columnEI++] = ListEI.TypeName;
                worksheetEI.Cells[rowEI, columnEI++] = ListEI.TypeNameId;
                worksheetEI.Cells[rowEI, columnEI++] = ListEI.FamilyName;
                worksheetEI.Cells[rowEI, columnEI++] = ListEI.Volume;
                worksheetEI.Cells[rowEI, columnEI++] = ListEI.Area;
                worksheetEI.Cells[rowEI, columnEI++] = ListEI.Length;
            }

            // Fitting
            var range_rowsEI = worksheetEI.get_Range("A1", "Z1").EntireColumn; // all cels
            range_rowsEI.AutoFit();

            // Color
            Excel.Range rangeA = (Excel.Range)worksheetEI.UsedRange.Columns["A:S"]; // all cells from column A to S
            foreach (Excel.Range cell in rangeA.Cells)
            {
                if (cell.Value != null && cell.Value.ToString() == "*NA*") cell.Font.Color = System.Drawing.Color.FromArgb(150, 150, 150); // Not applicable in grey
            }

            #endregion OTHER ROWS - Instance Elements
        }

        public static void CreateElementTypesSpreadsheet(List<ET> ETSSorted)
        {
            #region Excel Sheet "Element Types"

            Excel.Worksheet worksheetET = (Excel.Worksheet)workbook.Worksheets.Add(); // adds a new worksheet
            worksheetET.Name = "Element Types"; // worksheetET Name

            #endregion Excel Sheet "Element Types"

            #region HEADER - 1ST ROW - ElementTypes

            int columnET = 1;
            int rowET = 1;

            worksheetET.Cells[rowET, columnET++] = "ID";
            worksheetET.Cells[rowET, columnET++] = "IsType";
            worksheetET.Cells[rowET, columnET++] = "Category Name";
            worksheetET.Cells[rowET, columnET++] = "Category Id";
            worksheetET.Cells[rowET, columnET++] = "Type Name";
            worksheetET.Cells[rowET, columnET++] = "Family Name";
            worksheetET.Cells[rowET, columnET++] = "Quantity";
            worksheetET.Cells[rowET, columnET++] = "Total Volume";
            worksheetET.Cells[rowET, columnET++] = "Total Area";
            worksheetET.Cells[rowET, columnET++] = "Total Length";
            worksheetET.Cells[rowET, columnET++] = "Cost per unit";
            worksheetET.Cells[rowET, columnET++] = "Assembly Code";
            worksheetET.Cells[rowET, columnET++] = "Assembly Description";
            worksheetET.Cells[rowET, columnET++] = "Keynote Value";
            worksheetET.Cells[rowET, columnET++] = "Keynote Text";

            // Fitting
            var range_rowET = worksheetET.get_Range("A1", "A1").EntireRow; // all cels of the first row
            range_rowET.Font.Bold = true;
            range_rowET.EntireColumn.AutoFit();

            #endregion HEADER - 1ST ROW - ElementTypes

            #region OTHER ROWS - ElementTypes

            //Filling
            foreach (var ListET in ETSSorted)
            {
                ++rowET; // +1 row
                columnET = 1;
                // writes each element of ListET in a cell fo each column
                worksheetET.Cells[rowET, columnET++] = ListET.ID; // column 1 - eT.Id.IntegerValue
                worksheetET.Cells[rowET, columnET++] = ListET.IsType; // column 2 - ElementType of eT
                worksheetET.Cells[rowET, columnET++] = ListET.CategoryName; // column 3 - Category.Name of eT
                worksheetET.Cells[rowET, columnET++] = ListET.CategoryId; // column 4 - Category of eT
                worksheetET.Cells[rowET, columnET++] = ListET.TypeName; // column 5 - Name of eT
                worksheetET.Cells[rowET, columnET++] = ListET.FamilyName; // column 6 - Family Name of eT
                worksheetET.Cells[rowET, columnET++] = ListET.Quantity; // column 7 - Quantity of each eF of eT
                if (ListET.TotalVolume != null) worksheetET.Cells[rowET, columnET++] = ListET.TotalVolume; // column 8 - Total Volume of eT
                else worksheetET.Cells[rowET, columnET++] = "*NA*";
                if (ListET.TotalArea != null) worksheetET.Cells[rowET, columnET++] = ListET.TotalArea; // column 9 - Total Area of eT
                else worksheetET.Cells[rowET, columnET++] = "*NA*";
                if (ListET.TotalLength != null) worksheetET.Cells[rowET, columnET++] = ListET.TotalLength; // column 10 - Total Length de eT
                else worksheetET.Cells[rowET, columnET++] = "*NA*";
                worksheetET.Cells[rowET, columnET++] = ListET.Cost; // column 11 - Total Cost of eT
                worksheetET.Cells[rowET, columnET++] = ListET.AssemblyCode; // column 18 - AssemblyCode of eT
                worksheetET.Cells[rowET, columnET++] = ListET.AssemblyDesc; // column 19 - AssemblyDesc of eT
                worksheetET.Cells[rowET, columnET++] = ListET.KeyValue; // column 20 - KeyValue of eT
                worksheetET.Cells[rowET, columnET++] = ListET.KeyText; // column 21 - KeyText of eT
            }

            // Fitting
            var range_rowsET = worksheetET.get_Range("A2:Z" + rowET);
            range_rowsET.EntireColumn.AutoFit();

            // Colors
            Excel.Range rangeB = (Excel.Range)worksheetET.UsedRange.Columns["A:Z"]; // all cells from column A to Z
            foreach (Excel.Range cell in rangeB.Cells)
            {
                if (cell.Value != null && cell.Value.ToString() == "*NA*")
                    cell.Font.Color = System.Drawing.Color.FromArgb(150, 150, 150); // Not applicable in grey
                if (cell.Value != null && cell.Value.ToString() == "MISSING")
                    cell.Font.Color = System.Drawing.Color.FromArgb(255, 0, 0); // Missing in red
            }

            #endregion OTHER ROWS - ElementTypes
        }

        public static void CreateBillofQuantitiesSpreadsheet(UIApplication uiapp, List<ET> ETS, string docTitle)
        {
            List<BQ> BQS = RevitUtils.RetrieveBQData(uiapp, ETS);

            #region Excel Sheet "Bill of Quantities"

            Excel.Worksheet worksheetBQ = (Excel.Worksheet)workbook.Worksheets.Add(); // adds a new worksheet
            worksheetBQ.Name = "Bill of Quantities"; // worksheetBQ Name

            #endregion // Launch or access Excel via COM Interop

            #region Header formatting

            //cells E1 to E10
            Excel.Range rangeE = worksheetBQ.Range["E1:E10"];
            rangeE.Font.Bold = true;
            rangeE.Font.Color = System.Drawing.Color.FromArgb(0, 51, 102);

            //merge of cells E1 to E10
            for (int row = 1; row <= 10; row++) worksheetBQ.Range[worksheetBQ.Cells[row, 5], worksheetBQ.Cells[row, 7]].Merge();

            //cells E2 and E3
            worksheetBQ.Cells[2, 5] = "REVIT MODEL DOCUMENT";
            worksheetBQ.Cells[3, 5] = docTitle;

            // cell E9
            worksheetBQ.Cells[9, 5] = "BILL OF QUANTITIES";

            // cell E10
            string monthyear = DateTime.Now.ToString("MMMM yyyy", CultureInfo.CurrentCulture).ToUpper();
            worksheetBQ.Cells[10, 5] = monthyear;

            #endregion // Header formatting

            #region Columns Width

            worksheetBQ.Columns["A:A"].ColumnWidth = 4.57;
            worksheetBQ.Columns["B:B"].ColumnWidth = 21.43;
            worksheetBQ.Columns["C:C"].ColumnWidth = 43.57;
            worksheetBQ.Columns["D:D"].ColumnWidth = 5.43;
            worksheetBQ.Columns["E:E"].ColumnWidth = 12.43;
            worksheetBQ.Columns["F:F"].ColumnWidth = 13.29;
            worksheetBQ.Columns["G:G"].ColumnWidth = 13.29;
            worksheetBQ.Columns["H:H"].ColumnWidth = 15.86;

            #endregion // Columns Width

            #region HEADER - 1ST ROW - "Bill of Quantities"

            // header cells merge
            worksheetBQ.Range["B12:B13"].Merge();
            worksheetBQ.Range["C12:C13"].Merge();
            worksheetBQ.Range["D12:D13"].Merge();
            worksheetBQ.Range["E12:E13"].Merge();
            worksheetBQ.Range["F12:F13"].Merge();
            worksheetBQ.Range["G12:H12"].Merge();

            // header cells names
            worksheetBQ.Cells[12, 2] = "Nº ART."; //B12
            worksheetBQ.Cells[12, 3] = "DESIGNATION"; //C12
            worksheetBQ.Cells[12, 4] = "UN"; //D12
            worksheetBQ.Cells[12, 5] = "QUANT."; //E12
            worksheetBQ.Cells[12, 6].Style.WrapText = true; // to visible the line brakes //F12
            worksheetBQ.Cells[12, 6] = "PRICE\n PER UNIT"; //F12
            worksheetBQ.Cells[12, 7] = "COST"; //G12
            worksheetBQ.Cells[13, 7] = "PARTIAL"; //G13
            worksheetBQ.Cells[13, 8] = "TOTAL"; //H13

            // header cells range
            Excel.Range rangeH = worksheetBQ.Range["B12:H13"];

            // header cells style
            rangeH.Style.Font.Name = "Arial";
            rangeH.Style.Font.Size = 10;
            rangeH.Font.Bold = true;

            // header cells centered and aligned
            rangeH.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            rangeH.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // header cells boundaries
            rangeH.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            #endregion // HEADER - 1ST ROW - "Bill of Quantities"

            #region OTHER ROWS - "Bill of Quantities"

            #region Table filling

            int rowBQ = 14;

            var BQSSorted = BQS.AsQueryable().OrderBy(bq => bq.KeyValue).ToList();

            foreach (var ListBQ in BQSSorted)
            {
                ++rowBQ; // +1 row
                worksheetBQ.Cells[rowBQ, 2] = ListBQ.KeyValue;
                if (ListBQ.KeyValue == "MISSING")
                {
                    Excel.Range range = worksheetBQ.Cells[rowBQ, 2];
                    range.Font.Color = System.Drawing.Color.FromArgb(255, 0, 0); // Missing KeyValue
                }

                Excel.Range range3 = worksheetBQ.Cells[rowBQ, 3];
                range3.Font.Bold = true;
                range3.Style.WrapText = true;
                range3.Value = ListBQ.AssemblyDesc;
                if (ListBQ.AssemblyDesc == "MISSING") range3.Font.Color = System.Drawing.Color.FromArgb(255, 0, 0); // Missing AssemblyDesc

                Excel.Range range4 = worksheetBQ.Cells[rowBQ + 1, 3];
                worksheetBQ.Cells[rowBQ + 1, 3] = ListBQ.KeyText;
                if (ListBQ.KeyText == "MISSING") range4.Font.Color = System.Drawing.Color.FromArgb(255, 0, 0); // Missing KeynoteText

                worksheetBQ.Cells[rowBQ, 4] = ListBQ.Unit;

                worksheetBQ.Range[worksheetBQ.Cells[rowBQ, 4], worksheetBQ.Cells[rowBQ + 1, 4]].Merge();
                Excel.Range range5 = worksheetBQ.Cells[rowBQ, 5]; // column E
                range5.NumberFormat = "###0.00"; // column E in euros
                range5.Value = ListBQ.Quant;
                worksheetBQ.Range[worksheetBQ.Cells[rowBQ, 5], worksheetBQ.Cells[rowBQ + 1, 5]].Merge();
                Excel.Range range6 = worksheetBQ.Cells[rowBQ, 6]; // column F
                range6.NumberFormat = "###0.00 €"; // column F in euros
                range6.Value = ListBQ.PrUnit;
                worksheetBQ.Range[worksheetBQ.Cells[rowBQ, 6], worksheetBQ.Cells[rowBQ + 1, 6]].Merge();
                Excel.Range range7 = worksheetBQ.Cells[rowBQ, 7]; // column G
                range7.NumberFormat = "###0.00 €"; // column G in euros
                range7.Value = ListBQ.Partial;
                worksheetBQ.Range[worksheetBQ.Cells[rowBQ, 7], worksheetBQ.Cells[rowBQ + 1, 7]].Merge();
                ++rowBQ; // +1 linha
                ++rowBQ; // +1 linha
            }
            Excel.Range range8 = worksheetBQ.Cells[15, 8]; // cell H15
            range8.NumberFormat = "###0.00 €"; // cell H15 in euros
            range8.Value = "=Sum(G15:G" + rowBQ + ")";
            worksheetBQ.Range[worksheetBQ.Cells[15, 8], worksheetBQ.Cells[16, 8]].Merge();

            #endregion // Table filling

            #region Table formatting

            // gridlines turned off
            excel.ActiveWindow.DisplayGridlines = false;
            // table aligned
            worksheetBQ.Range["B15:C" + rowBQ].VerticalAlignment = Excel.XlHAlign.xlHAlignGeneral;

            worksheetBQ.Range["D15:H" + rowBQ].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            worksheetBQ.Range["D15:H" + rowBQ].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // table left boarders
            worksheetBQ.Range["B14:B" + rowBQ].Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
            worksheetBQ.Range["D14:D" + rowBQ].Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
            worksheetBQ.Range["F14:F" + rowBQ].Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
            worksheetBQ.Range["G14:G" + rowBQ].Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
            worksheetBQ.Range["I14:I" + rowBQ].Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;

            // table bottom boarders
            worksheetBQ.Range["B" + rowBQ + ":H" + rowBQ].Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;

            // Columns B and C text aligned to the left
            Excel.Range range9 = worksheetBQ.Range["B:C"];
            range9.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            #endregion // Table formatting

            #endregion // OTHER ROWS - "Bill of Quantities"
        }
        public static void EnableInteraction()
        {
            excel.Interactive = true; //allows the user to interact with Excel
        }
        
    }
}
