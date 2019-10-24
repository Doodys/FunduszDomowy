using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;

namespace FunduszDomowy
{
    class GemBoxExcel
    {
        static Config oConfig = new Config();
        static string sExcelFileDir = oConfig.SaveDir;
        static DateTime dtCurrentDateTime = DateTime.Now;
        string sWorksheetName = string.Format(dtCurrentDateTime.ToString("MM-yyyy", new CultureInfo("pl-PL")));
        public void CreateSpreadsheet()
        {
            try
            {
                var workbook = ExcelFile.Load(sExcelFileDir + "Fundusz_Domowy.xlsx");
                ExcelWorksheet oWorksheet = workbook.Worksheets.ActiveWorksheet;

                for (int i = 0; i <= workbook.Worksheets.Count; i++)
                {
                    if (workbook.Worksheets[i].Name.Equals(sWorksheetName))
                    {
                        oWorksheet = workbook.Worksheets[i];
                        break;
                    }
                }

                if (oWorksheet.Cells["B2"].Value == null || oWorksheet.Cells["C2"].Value == null || oWorksheet.Cells["D2"].Value == null || oWorksheet.Cells["F2"].Value == null)
                {
                    oWorksheet.Cells["B2"].Value = "Kwota";
                    oWorksheet.Cells["C2"].Value = "Data";
                    oWorksheet.Cells["D2"].Value = "Godzina";
                    oWorksheet.Cells["F2"].Value = "Suma wydatkow: ";
                }

                if (oWorksheet.Cells["G2"].Value == null && oConfig.ExcelLang.Equals("PL"))
                {
                    oWorksheet.Cells["G2"].Formula = "=SUMA(B:B)";
                }
                else if (oWorksheet.Cells["G2"].Value == null && !oConfig.ExcelLang.Equals("PL"))
                {
                    oWorksheet.Cells["G2"].Formula = "=SUM(B:B)";
                }

                //columns' width
                oWorksheet.Columns["B"].SetWidth(15, LengthUnit.ZeroCharacterWidth);
                oWorksheet.Columns["C"].SetWidth(15, LengthUnit.ZeroCharacterWidth);
                oWorksheet.Columns["D"].SetWidth(15, LengthUnit.ZeroCharacterWidth);
                oWorksheet.Columns["F"].SetWidth(18, LengthUnit.ZeroCharacterWidth);
                oWorksheet.Columns["G"].SetWidth(15, LengthUnit.ZeroCharacterWidth);

                //headers' style
                var style = new CellStyle();
                style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                style.VerticalAlignment = VerticalAlignmentStyle.Center;
                style.FillPattern.SetSolid(Color.CornflowerBlue);
                style.Font.Weight = ExcelFont.BoldWeight;
                style.Font.Color = Color.White;

                oWorksheet.Cells["B2"].Style = style;
                oWorksheet.Cells["C2"].Style = style;
                oWorksheet.Cells["D2"].Style = style;
                oWorksheet.Cells["F2"].Style = style;
                oWorksheet.Cells["G2"].Style = style;

                workbook.Save(sExcelFileDir + "Fundusz_Domowy.xlsx");
            }
            catch (System.IO.FileNotFoundException)
            {
                var workbook = new ExcelFile();
                workbook.Worksheets.Add(sWorksheetName);
                workbook.Save(sExcelFileDir + "Fundusz_Domowy.xlsx");
                CreateSpreadsheet();
            }
            catch (ArgumentOutOfRangeException)
            {
                var workbook = ExcelFile.Load(sExcelFileDir + "Fundusz_Domowy.xlsx");
                workbook.Worksheets.Add(sWorksheetName);
                workbook.Save(sExcelFileDir + "Fundusz_Domowy.xlsx");
                CreateSpreadsheet();
            }
        }

        public void AddToSpreadsheet(List<string> lData)
        {
            var workbook = ExcelFile.Load(sExcelFileDir + "Fundusz_Domowy.xlsx");
            string sCellStart = "", sCellEnd = "";

            ExcelWorksheet oWorksheet = workbook.Worksheets.ActiveWorksheet;

            for (int i = 0; i <= workbook.Worksheets.Count; i++)
            {
                if (workbook.Worksheets[i].Name.Equals(sWorksheetName))
                {
                    oWorksheet = workbook.Worksheets[i];
                    break;
                }
            }
            List<long> lList = new List<long>();
            foreach (var row in oWorksheet.Rows)
            {
                foreach (var cell in row.AllocatedCells)
                {
                    lList.Add(Convert.ToInt64(cell.Name.Substring(1)));
                }
            }

            long iLast = lList[lList.Count - 1];
            iLast++;
            oWorksheet.Cells["B" + iLast].SetValue(Convert.ToDouble(lData[0]));
            oWorksheet.Cells["C" + iLast].Value = lData[1];
            oWorksheet.Cells["D" + iLast].Value = lData[2];

            sCellStart = "B" + iLast.ToString();
            sCellEnd = "D" + iLast.ToString();

            var style = new CellStyle();
            style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            style.VerticalAlignment = VerticalAlignmentStyle.Center;

            oWorksheet.Cells.GetSubrange(sCellStart + ":" + sCellEnd).Style = style;
            oWorksheet.Cells.GetSubrange(sCellStart + ":" + sCellEnd).Style.Borders.SetBorders(MultipleBorders.All, Color.Black, LineStyle.Thin);

            workbook.Save(sExcelFileDir + "Fundusz_Domowy.xlsx");
            lList.Clear();
        }
    }
}
