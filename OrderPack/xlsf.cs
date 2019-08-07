using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;

namespace XlsFile
{
    public class xlsf
    {

        #region App
        public Excel.Application excelApp;

        public void Quit()
        {
            excelApp.Quit();
        }

        public void SetVisible(bool IsVisible)
        {
            excelApp.Visible = IsVisible; //false, 速度快
        }
        #endregion

        //exception

        #region constructor, destructor
        public xlsf()
        {
            excelApp = new Excel.Application();
        }

        public xlsf(object AppProcess)
        {
            //in main.cs
            //xlsf excel_file = new xlsf(Marshal.GetActiveObject("Excel.Application"));
            excelApp = AppProcess as Excel.Application;
        }

        ~xlsf()
        {
            Marshal.FinalReleaseComObject(excelApp);
        }
        #endregion

        #region Workbooks, Workbook
        //Excel._Worksheet objSheet;
        private Excel.Workbooks CurrWorkbooks
        {
            get { return excelApp.Workbooks; }
        }

        private Excel.Workbook CurrWorkbook
        {
            get { return excelApp.ActiveWorkbook; }
        }

        public void NewFile()
        {
            CurrWorkbooks.Add();
        }

        public void OpenFile(string XlsFilePathName)
        {
            CurrWorkbooks.Open(XlsFilePathName);
        }

        public void CloseFile(bool IsCheckSaveFile = true)
        {
            CurrWorkbook.Close(IsCheckSaveFile);
        }

        public void Save()
        {
            CurrWorkbook.Save();
        }

        public void SaveAs(string FilePathName)
        {
            excelApp.DisplayAlerts = false;
            CurrWorkbook.SaveAs(FilePathName);
            excelApp.DisplayAlerts = true;
        }
        #endregion

        #region Sheet
        private Excel.Sheets CurrSheets
        {
            get { return excelApp.Sheets; }
        }

        private Excel._Worksheet CurrSheet
        {
            get { return excelApp.ActiveSheet; }
        }

        public void NewSheet()
        {
            CurrSheets.Add();
        }

        public void CopySheet(long NewSheetIndex = -1)
        {
            if (CheckSheetIndex(NewSheetIndex))
                NewSheetIndex = SheetTotal();

            CurrSheet.Copy(CurrSheets[NewSheetIndex]);
        }

        public void DeleteSheet()
        {
            excelApp.DisplayAlerts = false;
            CurrSheet.Delete();
            excelApp.DisplayAlerts = true;
        }

        public xlsf SelectSheet(int SheetIndex)
        {
            CurrSheets[SheetIndex].Select();
            return this;
        }

        public xlsf SelectSheet(string SheetName)
        {
            CurrSheets[SheetName].Select();
            return this;
        }

        public void PrintSheet(bool IsPrintPreview)
        {
            if (IsPrintPreview)
                CurrSheet.PrintPreview();
            else
                CurrSheet.PrintOut(1, 1, 2, true);
        }

        public void MoveSheet(long SheetIndex = -1)
        {
            if ( CheckSheetIndex(SheetIndex))
                SheetIndex = SheetTotal();

            CurrSheet.Move(CurrSheets[SheetIndex]);
        }

        private bool CheckSheetIndex(long SheetIndex)
        {
            return SheetIndex == -1 || (SheetIndex >= 0 && SheetIndex <= SheetTotal());
        }

        public void HiddenSheet(bool IsHide)
        {
            if (IsHide)
                CurrSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            else
                CurrSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
        }

        //由SheetNumber 取得SheetName
        public string GetSheetName()
        {
            return CurrSheet.Name;
        }

        public void SetSheetName(string SheetName)
        {
            CurrSheet.Name = SheetName;
        }

        public long SheetTotal()
        {
            return CurrSheets.Count;
        }
        #endregion

        #region Cell
        //Excel.Range m_Range;, m_Col, m_Row;
        private Excel.Range CurrCell
        {
            get { return excelApp.ActiveCell; }
        }

        #region AutoFit
        private Excel.Range CurrColumnCells
        {
            get { return CurrCell.EntireColumn; }
        }

        private Excel.Range CurrRowCells
        {
            get { return CurrCell.EntireRow; }
        }

        public void AutoFitWidth()
        {
            CurrColumnCells.AutoFit();
        }

        public void AutoFitHight()
        {
            CurrRowCells.AutoFit();
        }
        #endregion

        #region SelectCell
        public xlsf SelectCell(string SelectRange) //("A3") or ("A1:B3")
        {
            CurrSheet.Range[SelectRange].Select();
            return this;
        }

        public xlsf SelectCell(string X, int Y) //  ("A", 3)
        {
            CurrSheet.Cells[Y, X].Select();
            return this;
        }

        public xlsf SelectCell(string CellPosition1, string CellPosition2) //  ("A1", "B3")
        {
            CurrSheet.get_Range(CellPosition1, CellPosition2).Select();
            return this;
        }

        public xlsf OffsetSelectCell(int OffsetRight, int OffsetDown)
        {
            CurrCell.Offset[OffsetDown, OffsetRight].Select();
            //get_Offset 是舊版語法
            return this;
        }
        #endregion

        public void SetCell(object CellValue)
        {
            CurrCell.Value = CellValue;
            //Value2是舊版語法
        }

        public void ClearCell()
        {
            CurrCell.Clear();
        }


        public void CopyCell(string CellPosition1, string CellPosition2) //  ("A1", "B3")
        {
            CurrCell.AutoFill(excelApp.get_Range(CellPosition1, CellPosition2), Excel.XlAutoFillType.xlFillCopy);
        }

        #region GetCellValue
        public DateTime GetCell2DateTime()
        {
            object value = CurrCell.Value2;
            DateTime dt = new DateTime(); ;

            if (value != null)
            {
                if (value is double)
                {
                    dt = DateTime.FromOADate((double)value);
                }
                else
                {
                    DateTime.TryParse((string)value, out dt);
                }
            }
            return dt;
        }

        public string GetCell2Str()
        {
            if (CurrCell.Value == null)
                return "";
            else
                return CurrCell.Value.ToString();
        }

        public long GetCell2Int()
        {
            if (CurrCell.Value == null)
                return 0;
            else
                return Convert.ToInt64(CurrCell.Value);
        }

        public double GetCell2Double()
        {
            if (CurrCell.Value == null)
                return 0.0;
            else
                return Convert.ToDouble(CurrCell.Value);
        }
        #endregion

        public void FreezeCell(bool IsFreeze = true)
        {
            excelApp.ActiveWindow.FreezePanes = IsFreeze;
        }

        public void SelectCellandSetMerge(string SelectRange)
        {
            //透過 CurrCell(ActiveCell)無法Merge
            //暫時找不到取出儲存格座標值的方法！QQ
            if (CurrSheet.Range[SelectRange].MergeCells)
                CurrSheet.Range[SelectRange].UnMerge();
            else
                CurrSheet.Range[SelectRange].Merge();
        }

        public long GetHorztlStartCell()
        {
            return CurrCell.Column;
        }
        public long GetVrticlStartCell()
        {   
            return CurrCell.Row;
        }
        public long GetHorztlTotalCell()
        {
            //就是使用的範圍
            return CurrSheet.UsedRange.Cells.Columns.Count;
        }
        
        public long GetVrticlTotalCell()
        {
            //就是使用的範圍
            return CurrSheet.UsedRange.Cells.Rows.Count;
        }

        #region 對齊
        public xlsf SetHorztlAlgmet(Excel.XlHAlign AlignType)
        {
            CurrCell.HorizontalAlignment = AlignType;
            return this;
        }
        public xlsf SetVrticlAlgmet(Excel.XlVAlign AlignType)
        {
            CurrCell.VerticalAlignment = AlignType;
            return this;
        }
        public xlsf SetTextAngle(int Angle)
        {
            CurrCell.Orientation = Angle;
            return this;
        }
        public xlsf AutoNewLine(bool NewLine = true)
        {
            CurrCell.WrapText = NewLine;
            return this;
        }
        #endregion

        public xlsf SetCellBorder(Excel.XlLineStyle BoarderStyle, Excel.XlBorderWeight BoarderWeight, Color BoarderColor)
        {
            CurrCell.Borders.LineStyle = BoarderStyle;
            CurrCell.Borders.Weight = BoarderWeight;
            CurrCell.Borders.Color = ColorTranslator.ToOle(BoarderColor);

            return this;
        }

        public xlsf SetCellColor(Color ColorObj, Excel.XlPattern PatternType = Excel.XlPattern.xlPatternAutomatic)
        {
            CurrCell.Interior.Color = ColorTranslator.ToOle(ColorObj);
            CurrCell.Interior.Pattern = PatternType;
            return this;
        }

        public xlsf SetCellBk(Color ColorObj, Excel.XlPattern PatternType = Excel.XlPattern.xlPatternAutomatic)
        {
            CurrCell.Interior.Color = ColorTranslator.ToOle(ColorObj);
            CurrCell.Interior.Pattern = PatternType;
            return this;
        }

        #region Font
        public xlsf SetFont(string FontName = "微軟正黑體")
        {
            //Excel.Style style = Globals.ThisWorkbook.Styles.Add("NewStyle");
            CurrCell.Font.Name = FontName;
            return this;
        }

        public xlsf SetFontSize(int FontSize = 12)
        {
            CurrCell.Font.Size = FontSize;
            return this;
        }

        public xlsf SetFontColor(Color FontColor)
        {
            CurrCell.Font.Color = System.Drawing.ColorTranslator.ToOle(FontColor);
            return this;
        }

        public xlsf SetFontBold(bool IsBold)
        {
            CurrCell.Font.Bold = IsBold;
            return this;
        }

        public xlsf SetFontStrkthrgh(bool IsStrkthrgh)
        {
            CurrCell.Font.Strikethrough = IsStrkthrgh;
            return this;
        }
        #endregion

        public xlsf SetCellHeight(int HeightValue)
        {
            CurrCell.RowHeight = HeightValue;
            return this;
        }

        public xlsf SetCellWidth(int WidthValue)
        {
            CurrCell.ColumnWidth = WidthValue;
            return this;
        }

        public xlsf Insert(Excel.XlInsertShiftDirection ShiftCellType)
        {
            CurrCell.Insert(ShiftCellType);
            return this;
        }
        #endregion
    }
}
