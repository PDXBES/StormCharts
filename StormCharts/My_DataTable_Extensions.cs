using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace StormCharts
{
  public static class My_DataTable_Extensions
  {
    // Export DataTable into an excel file with field names in the header line
    // - Save excel file without ever making it visible if filepath is given
    // - Don't save excel file, just make it visible if no filepath is given
    public static void ExportToExcel(this DataTable Tbl, string shortDate, DateTime theStart, string ExcelFilePath = null, int WorkSheetName = 0, int WorkSheetCount = 0, bool LastStorm = false)
    {
      // load excel, and create a new workbook
      Excel.Application excelApp = new Excel.Application();
      Excel.Workbook wb = null;
      Excel._Worksheet workSheet;

      try
      {
        int FileNew = 0;

        if (Tbl == null || Tbl.Columns.Count == 0)
          throw new Exception("ExportToExcel: Null or empty input table!\n");


        if (WorkSheetCount == 1)
        {
          wb = excelApp.Workbooks.Add(Type.Missing);
          workSheet = excelApp.ActiveSheet;
          workSheet.Name = "All Gauges";
          workSheet.get_Range("AA1").Value = "Storm Period, hours";
          workSheet.get_Range("AB1").Value = "1";
          workSheet.get_Range("AC1").Value = "2";
          workSheet.get_Range("AD1").Value = "3";
          workSheet.get_Range("AE1").Value = "6";
          workSheet.get_Range("AF1").Value = "12";
          workSheet.get_Range("AG1").Value = "24";
          workSheet.get_Range("AH1").Value = "48";
          workSheet.get_Range("AI1").Value = "72";

          workSheet.get_Range("A2").Value = "Hours";
          workSheet.get_Range("D2").Value = "1";
          workSheet.get_Range("E2").Value = "2";
          workSheet.get_Range("F2").Value = "3";
          workSheet.get_Range("G2").Value = "6";
          workSheet.get_Range("H2").Value = "12";
          workSheet.get_Range("I2").Value = "24";
          workSheet.get_Range("J2").Value = "48";
          workSheet.get_Range("A3").Value = "ASFO Citywide 4-Per-Winter";
          workSheet.get_Range("D3").Value = "0.24";
          workSheet.get_Range("E3").Value = "0.34";
          workSheet.get_Range("F3").Value = "0.44";
          workSheet.get_Range("G3").Value = "0.65";
          workSheet.get_Range("H3").Value = "0.89";
          workSheet.get_Range("I3").Value = "1.19";
          workSheet.get_Range("J3").Value = "1.53";
          workSheet.get_Range("A4").Value = "3-per-Winter Storm";
          workSheet.get_Range("D4").Value = "0.27";
          workSheet.get_Range("E4").Value = "0.38";
          workSheet.get_Range("F4").Value = "0.49";
          workSheet.get_Range("G4").Value = "0.72";
          workSheet.get_Range("H4").Value = "1.01";
          workSheet.get_Range("I4").Value = "1.35";
          workSheet.get_Range("J4").Value = "1.74";
          workSheet.get_Range("A5").Value = "2-per-Winter Storm";
          workSheet.get_Range("D5").Value = "0.30";
          workSheet.get_Range("E5").Value = "0.43";
          workSheet.get_Range("F5").Value = "0.55";
          workSheet.get_Range("G5").Value = "0.81";
          workSheet.get_Range("H5").Value = "1.17";
          workSheet.get_Range("I5").Value = "1.59";
          workSheet.get_Range("J5").Value = "2.07";
          workSheet.get_Range("A6").Value = "1-per-Winter Storm";
          workSheet.get_Range("D6").Value = "0.35";
          workSheet.get_Range("E6").Value = "0.50";
          workSheet.get_Range("F6").Value = "0.65";
          workSheet.get_Range("G6").Value = "0.97";
          workSheet.get_Range("H6").Value = "1.43";
          workSheet.get_Range("I6").Value = "1.93";
          workSheet.get_Range("J6").Value = "2.55";
          workSheet.get_Range("A7").Value = "ASFO 5-Year Winter";
          workSheet.get_Range("D7").Value = "0.43";
          workSheet.get_Range("E7").Value = "0.62";
          workSheet.get_Range("F7").Value = "0.80";
          workSheet.get_Range("G7").Value = "1.21";
          workSheet.get_Range("H7").Value = "1.81";
          workSheet.get_Range("I7").Value = "2.51";
          workSheet.get_Range("J7").Value = "3.26";
          workSheet.get_Range("A8").Value = "BES Sewer Design Manual - Design Storms";
          workSheet.get_Range("A9").Value = "BES 2-Year Storm";
          workSheet.get_Range("D9").Value = "0.46";
          workSheet.get_Range("E9").Value = "0.64";
          workSheet.get_Range("F9").Value = "0.80";
          workSheet.get_Range("G9").Value = "1.19";
          workSheet.get_Range("H9").Value = "1.78";
          workSheet.get_Range("I9").Value = "2.40";
          workSheet.get_Range("A10").Value = "BES 5-Year Storm";
          workSheet.get_Range("D10").Value = "0.59";
          workSheet.get_Range("E10").Value = "0.80";
          workSheet.get_Range("F10").Value = "0.99";
          workSheet.get_Range("G10").Value = "1.49";
          workSheet.get_Range("H10").Value = "2.18";
          workSheet.get_Range("I10").Value = "2.93";
          workSheet.get_Range("A11").Value = "BES 10-Year Storm";
          workSheet.get_Range("D11").Value = "0.68";
          workSheet.get_Range("E11").Value = "0.92";
          workSheet.get_Range("F11").Value = "1.15";
          workSheet.get_Range("G11").Value = "1.68";
          workSheet.get_Range("H11").Value = "2.45";
          workSheet.get_Range("I11").Value = "3.34";
          workSheet.get_Range("A12").Value = "BES 25-Year Storm";
          workSheet.get_Range("D12").Value = "0.79";
          workSheet.get_Range("E12").Value = "1.06";
          workSheet.get_Range("F12").Value = "1.30";
          workSheet.get_Range("G12").Value = "1.91";
          workSheet.get_Range("H12").Value = "2.81";
          workSheet.get_Range("I12").Value = "3.77";
          workSheet.get_Range("A13").Value = "BES 50-Year Storm";
          workSheet.get_Range("D13").Value = "0.90";
          workSheet.get_Range("E13").Value = "1.18";
          workSheet.get_Range("F13").Value = "1.43";
          workSheet.get_Range("G13").Value = "2.13";
          workSheet.get_Range("H13").Value = "3.14";
          workSheet.get_Range("I13").Value = "4.20";
          workSheet.get_Range("A14").Value = "BES 100-Year Storm";
          workSheet.get_Range("D14").Value = "0.99";
          workSheet.get_Range("E14").Value = "1.30";
          workSheet.get_Range("F14").Value = "1.59";
          workSheet.get_Range("G14").Value = "2.34";
          workSheet.get_Range("H14").Value = "3.42";
          workSheet.get_Range("I14").Value = "4.61";
          workSheet.get_Range("A15").Value = "3-Year Summer Storm";
          workSheet.get_Range("D15").Value = "0.40";
          workSheet.get_Range("E15").Value = "0.52";
          workSheet.get_Range("F15").Value = "0.60";
          workSheet.get_Range("G15").Value = "0.85";
          workSheet.get_Range("H15").Value = "1.10";
          workSheet.get_Range("I15").Value = "1.41";
          workSheet.get_Range("J15").Value = "2.12";
          workSheet.get_Range("A16").Value = "10-Year Summer Storm";
          workSheet.get_Range("D16").Value = "0.51";
          workSheet.get_Range("E16").Value = "0.70";
          workSheet.get_Range("F16").Value = "0.85";
          workSheet.get_Range("G16").Value = "1.25";
          workSheet.get_Range("H16").Value = "1.68";
          workSheet.get_Range("I16").Value = "2.06";
          workSheet.get_Range("J16").Value = "3.15";



          wb.SaveAs(ExcelFilePath);
          //excelApp.Quit();
          //excelApp = new Excel.Application();
        }

        try
        {
          wb = excelApp.Workbooks.Open(ExcelFilePath);
          workSheet = excelApp.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
          FileNew = 0;
        }
        catch (Exception ex)
        {
          FileNew = 1;
          excelApp.Workbooks.Add();
          workSheet = excelApp.ActiveSheet;
        }

        // single worksheet
        workSheet.Name = WorkSheetName.ToString();// +"_" + shortDate.Replace('/', '_');

        // column headings
        for (int i = 0; i < Tbl.Columns.Count; i++)
        {
          workSheet.Cells[1, (i + 1)] = Tbl.Columns[i].ColumnName;
        }

        // rows
        for (int i = 0; i < Tbl.Rows.Count; i++)
        {
          // to do: format datetime values before printing
          for (int j = 0; j < Tbl.Columns.Count; j++)
          {
            workSheet.Cells[(i + 2), (j + 1)] = Tbl.Rows[i][j];
          }
        }

        //Create chart
        Excel.Range chartRange;

        Excel.ChartObjects xlCharts = (Excel.ChartObjects)workSheet.ChartObjects(Type.Missing);
        Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
        Excel.Chart chartPage = myChart.Chart;

        chartRange = workSheet.get_Range("D:D, F:F", Type.Missing);//"D2:D120, E2:E120"

        chartPage.SetSourceData(chartRange, Type.Missing);
        chartPage.ChartType = Excel.XlChartType.xlXYScatterLines;
        chartPage.Location(Excel.XlChartLocation.xlLocationAsNewSheet, WorkSheetName.ToString() + "_Chart");

        Excel.Range sumRange;

        if (workSheet.get_Range("F13").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("G13");
          sumRange.Formula = "=IF(ISBLANK(F2),TRUE,SUM(F2:F13))";
          try
          {
            sumRange.AutoFill(workSheet.get_Range("G13", "G" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);
          }
          catch (Exception ex)
          {
            //Autofill only works when applied to more than one cell.
          }
          sumRange = workSheet.get_Range("G1");
          sumRange.Formula = "=MAX(G13:G" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("F25").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("H25");
          sumRange.Formula = "=IF(ISBLANK(F2),TRUE,SUM(F2:F25))";
          sumRange.AutoFill(workSheet.get_Range("H25", "H" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("H1");
          sumRange.Formula = "=MAX(H25:H" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("F37").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("I37");
          sumRange.Formula = "=IF(ISBLANK(F2),TRUE,SUM(F2:F37))";
          sumRange.AutoFill(workSheet.get_Range("I37", "I" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("I1");
          sumRange.Formula = "=MAX(I37:I" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("F73").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("J73");
          sumRange.Formula = "=IF(ISBLANK(F2),TRUE,SUM(F2:F73))";
          sumRange.AutoFill(workSheet.get_Range("J73", "J" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("J1");
          sumRange.Formula = "=MAX(J73:J" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("F145").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("K145");
          sumRange.Formula = "=IF(ISBLANK(F2),TRUE,SUM(F2:F145))";
          sumRange.AutoFill(workSheet.get_Range("K145", "K" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("K1");
          sumRange.Formula = "=MAX(K145:K" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("F289").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("L289");
          sumRange.Formula = "=IF(ISBLANK(F2),TRUE,SUM(F2:F289))";
          sumRange.AutoFill(workSheet.get_Range("L289", "L" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("L1");
          sumRange.Formula = "=MAX(L289:L" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("F577").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("M577");
          sumRange.Formula = "=IF(ISBLANK(F2),TRUE,SUM(F2:F577))";
          sumRange.AutoFill(workSheet.get_Range("M577", "M" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("M1");
          sumRange.Formula = "=MAX(M577:M" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("F865").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("N865");
          sumRange.Formula = "=IF(ISBLANK(F2),TRUE,SUM(F2:F865))";
          sumRange.AutoFill(workSheet.get_Range("N865", "N" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("N1");
          sumRange.Formula = "=MAX(N865:N" + (Tbl.Rows.Count + 1).ToString();
        }

        ((Excel.Worksheet)wb.Sheets["All Gauges"]).get_Range("AA" + (WorkSheetName + 1).ToString()).Value = WorkSheetName.ToString();
        ((Excel.Worksheet)wb.Sheets["All Gauges"]).get_Range("AB" + (WorkSheetName + 1).ToString()).Value = workSheet.get_Range("F1").Value;
        ((Excel.Worksheet)wb.Sheets["All Gauges"]).get_Range("AC" + (WorkSheetName + 1).ToString()).Value = workSheet.get_Range("G1").Value;
        ((Excel.Worksheet)wb.Sheets["All Gauges"]).get_Range("AD" + (WorkSheetName + 1).ToString()).Value = workSheet.get_Range("H1").Value;
        ((Excel.Worksheet)wb.Sheets["All Gauges"]).get_Range("AE" + (WorkSheetName + 1).ToString()).Value = workSheet.get_Range("I1").Value;
        ((Excel.Worksheet)wb.Sheets["All Gauges"]).get_Range("AF" + (WorkSheetName + 1).ToString()).Value = workSheet.get_Range("J1").Value;
        ((Excel.Worksheet)wb.Sheets["All Gauges"]).get_Range("AG" + (WorkSheetName + 1).ToString()).Value = workSheet.get_Range("K1").Value;
        ((Excel.Worksheet)wb.Sheets["All Gauges"]).get_Range("AH" + (WorkSheetName + 1).ToString()).Value = workSheet.get_Range("L1").Value;
        ((Excel.Worksheet)wb.Sheets["All Gauges"]).get_Range("AI" + (WorkSheetName + 1).ToString()).Value = workSheet.get_Range("M1").Value;

        // Create summary chart
        //Create chart
        //Excel.ChartObjects xlCharts = (Excel.ChartObjects)workSheet.ChartObjects(Type.Missing);
        Excel.ChartObject myChart2 = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
        Excel.Chart chartPage2 = myChart2.Chart;

        //chartRange = workSheet.get_Range("F1:M1, 'All Gauges'!D2:J2", Type.Missing);//"D2:D120, E2:E120"
        if (theStart.Month >= 5 && theStart.Month < 11)
        {
          chartRange = ((Excel.Worksheet)wb.Sheets["All Gauges"]).get_Range("D15:I15, AB" + (WorkSheetName + 1).ToString() + ":AI" + (WorkSheetName + 1).ToString());
        }
        else
        {
          chartRange = ((Excel.Worksheet)wb.Sheets["All Gauges"]).get_Range("D3:J3, AB" + (WorkSheetName + 1).ToString() + ":AI" + (WorkSheetName + 1).ToString());
        }

        chartPage2.SetSourceData(chartRange, Type.Missing);

        if (theStart.Month >= 5 && theStart.Month < 11)
        {
          ((Excel.Series)chartPage2.SeriesCollection(1)).Name = "3-year Summer Storm";
          //((Excel.Series)chartPage2.SeriesCollection(1)). = "3-year Summer Storm";
          ((Excel.Series)chartPage2.SeriesCollection(2)).Name = "Event Rainfall";
        }
        else
        {
          ((Excel.Series)chartPage2.SeriesCollection(1)).Name = "ASFO Citywide 4-Per-Winter";
          ((Excel.Series)chartPage2.SeriesCollection(2)).Name = "Event Rainfall";
        }

        chartPage2.ChartType = Excel.XlChartType.xlXYScatterLines;
        chartPage2.Location(Excel.XlChartLocation.xlLocationAsNewSheet, WorkSheetName.ToString() + "_SummaryChart");

        if (LastStorm == true)
        {
          workSheet = (Excel.Worksheet)wb.Worksheets["All Gauges"];

          Excel.Range MRange;
          //Table 2
          MRange = workSheet.get_Range("A32");
          MRange.Value = "Rainfall Compared Against Design Storms";
          MRange = workSheet.get_Range("F33");
          MRange.Value = "Maximum Rainfall Depths per Duration (inches)";
          MRange = workSheet.get_Range("D34");
          MRange.Value = "Days";
          MRange = workSheet.get_Range("F34");
          MRange.Value = "0.04";
          MRange = workSheet.get_Range("G34");
          MRange.Value = "0.08";
          MRange = workSheet.get_Range("H34");
          MRange.Value = "0.13";
          MRange = workSheet.get_Range("I34");
          MRange.Value = "0.25";
          MRange = workSheet.get_Range("J34");
          MRange.Value = "0.5";
          MRange = workSheet.get_Range("K34");
          MRange.Value = "1";
          MRange = workSheet.get_Range("L34");
          MRange.Value = "2";
          MRange = workSheet.get_Range("A35");
          MRange.Value = "Gage #";
          MRange = workSheet.get_Range("B35");
          MRange.Value = "HYDRA Rain Gauge Name";
          MRange = workSheet.get_Range("C35");
          MRange.Value = "Areal Weight";
          MRange = workSheet.get_Range("D35");
          MRange.Value = "Hours";
          MRange = workSheet.get_Range("F35");
          MRange.Value = "1";
          MRange = workSheet.get_Range("G35");
          MRange.Value = "2";
          MRange = workSheet.get_Range("H35");
          MRange.Value = "3";
          MRange = workSheet.get_Range("I35");
          MRange.Value = "6";
          MRange = workSheet.get_Range("J35");
          MRange.Value = "12";
          MRange = workSheet.get_Range("K35");
          MRange.Value = "24";
          MRange = workSheet.get_Range("L35");
          MRange.Value = "48";

          MRange = workSheet.get_Range("A36");
          MRange.Value = "174";
          MRange = workSheet.get_Range("B36");
          MRange.Value = "Arleta School";
          MRange = workSheet.get_Range("C36");
          MRange.Value = "10.4886774261159%";
          MRange = workSheet.get_Range("F36");
          MRange.Value = "='174'!G1";
          MRange = workSheet.get_Range("G36");
          MRange.Value = "='174'!H1";
          MRange = workSheet.get_Range("H36");
          MRange.Value = "='174'!I1";
          MRange = workSheet.get_Range("I36");
          MRange.Value = "='174'!J1";
          MRange = workSheet.get_Range("J36");
          MRange.Value = "='174'!K1";
          MRange = workSheet.get_Range("K36");
          MRange.Value = "='174'!L1";
          MRange = workSheet.get_Range("L36");
          MRange.Value = "='174'!M1";

          MRange = workSheet.get_Range("A37");
          MRange.Value = "64";
          MRange = workSheet.get_Range("B37");
          MRange.Value = "Harney Pump Station";
          MRange = workSheet.get_Range("C37");
          MRange.Value = "7.85195398684396%";
          MRange = workSheet.get_Range("F37");
          MRange.Value = "='64'!G1";
          MRange = workSheet.get_Range("G37");
          MRange.Value = "='64'!H1";
          MRange = workSheet.get_Range("H37");
          MRange.Value = "='64'!I1";
          MRange = workSheet.get_Range("I37");
          MRange.Value = "='64'!J1";
          MRange = workSheet.get_Range("J37");
          MRange.Value = "='64'!K1";
          MRange = workSheet.get_Range("K37");
          MRange.Value = "='64'!L1";
          MRange = workSheet.get_Range("L37");
          MRange.Value = "='64'!M1";

          MRange = workSheet.get_Range("A38");
          MRange.Value = "171";
          MRange = workSheet.get_Range("B38");
          MRange.Value = "Sunnyside School";
          MRange = workSheet.get_Range("C38");
          MRange.Value = "7.61297615898003%";
          MRange = workSheet.get_Range("F38");
          MRange.Value = "='171'!G1";
          MRange = workSheet.get_Range("G38");
          MRange.Value = "='171'!H1";
          MRange = workSheet.get_Range("H38");
          MRange.Value = "='171'!I1";
          MRange = workSheet.get_Range("I38");
          MRange.Value = "='171'!J1";
          MRange = workSheet.get_Range("J38");
          MRange.Value = "='171'!K1";
          MRange = workSheet.get_Range("K38");
          MRange.Value = "='171'!L1";
          MRange = workSheet.get_Range("L38");
          MRange.Value = "='171'!M1";

          MRange = workSheet.get_Range("A39");
          MRange.Value = "6";
          MRange = workSheet.get_Range("B39");
          MRange.Value = "Mt. Tabor Yard";
          MRange = workSheet.get_Range("C39");
          MRange.Value = "7.18518519105569%";
          MRange = workSheet.get_Range("F39");
          MRange.Value = "='6'!G1";
          MRange = workSheet.get_Range("G39");
          MRange.Value = "='6'!H1";
          MRange = workSheet.get_Range("H39");
          MRange.Value = "='6'!I1";
          MRange = workSheet.get_Range("I39");
          MRange.Value = "='6'!J1";
          MRange = workSheet.get_Range("J39");
          MRange.Value = "='6'!K1";
          MRange = workSheet.get_Range("K39");
          MRange.Value = "='6'!L1";
          MRange = workSheet.get_Range("L39");
          MRange.Value = "='6'!M1";

          MRange = workSheet.get_Range("A40");
          MRange.Value = "214";
          MRange = workSheet.get_Range("B40");
          MRange.Value = "OPB Office";
          MRange = workSheet.get_Range("C40");
          MRange.Value = "7.13101575926787%";
          MRange = workSheet.get_Range("F40");
          MRange.Value = "='214'!G1";
          MRange = workSheet.get_Range("G40");
          MRange.Value = "='214'!H1";
          MRange = workSheet.get_Range("H40");
          MRange.Value = "='214'!I1";
          MRange = workSheet.get_Range("I40");
          MRange.Value = "='214'!J1";
          MRange = workSheet.get_Range("J40");
          MRange.Value = "='214'!K1";
          MRange = workSheet.get_Range("K40");
          MRange.Value = "='214'!L1";
          MRange = workSheet.get_Range("L40");
          MRange.Value = "='214'!M1";

          MRange = workSheet.get_Range("A41");
          MRange.Value = "12";
          MRange = workSheet.get_Range("B41");
          MRange.Value = "Fernwood School";
          MRange = workSheet.get_Range("C41");
          MRange.Value = "6.71372143610241%";
          MRange = workSheet.get_Range("F41");
          MRange.Value = "='12'!G1";
          MRange = workSheet.get_Range("G41");
          MRange.Value = "='12'!H1";
          MRange = workSheet.get_Range("H41");
          MRange.Value = "='12'!I1";
          MRange = workSheet.get_Range("I41");
          MRange.Value = "='12'!J1";
          MRange = workSheet.get_Range("J41");
          MRange.Value = "='12'!K1";
          MRange = workSheet.get_Range("K41");
          MRange.Value = "='12'!L1";
          MRange = workSheet.get_Range("L41");
          MRange.Value = "='12'!M1";

          MRange = workSheet.get_Range("A42");
          MRange.Value = "181";
          MRange = workSheet.get_Range("B42");
          MRange.Value = "Multnomah Raingage";
          MRange = workSheet.get_Range("C42");
          MRange.Value = "6.69850340896201%";
          MRange = workSheet.get_Range("F42");
          MRange.Value = "='181'!G1";
          MRange = workSheet.get_Range("G42");
          MRange.Value = "='181'!H1";
          MRange = workSheet.get_Range("H42");
          MRange.Value = "='181'!I1";
          MRange = workSheet.get_Range("I42");
          MRange.Value = "='181'!J1";
          MRange = workSheet.get_Range("J42");
          MRange.Value = "='181'!K1";
          MRange = workSheet.get_Range("K42");
          MRange.Value = "='181'!L1";
          MRange = workSheet.get_Range("L42");
          MRange.Value = "='181'!M1";

          MRange = workSheet.get_Range("A43");
          MRange.Value = "164";
          MRange = workSheet.get_Range("B43");
          MRange.Value = "Ecoroof - SW 12th";
          MRange = workSheet.get_Range("C43");
          MRange.Value = "5.87171347460055%";
          MRange = workSheet.get_Range("F43");
          MRange.Value = "='164'!G1";
          MRange = workSheet.get_Range("G43");
          MRange.Value = "='164'!H1";
          MRange = workSheet.get_Range("H43");
          MRange.Value = "='164'!I1";
          MRange = workSheet.get_Range("I43");
          MRange.Value = "='164'!J1";
          MRange = workSheet.get_Range("J43");
          MRange.Value = "='164'!K1";
          MRange = workSheet.get_Range("K43");
          MRange.Value = "='164'!L1";
          MRange = workSheet.get_Range("L43");
          MRange.Value = "='164'!M1";

          MRange = workSheet.get_Range("A44");
          MRange.Value = "175";
          MRange = workSheet.get_Range("B44");
          MRange.Value = "Glencoe School";
          MRange = workSheet.get_Range("C44");
          MRange.Value = "5.50441521843925%";
          MRange = workSheet.get_Range("F44");
          MRange.Value = "='175'!G1";
          MRange = workSheet.get_Range("G44");
          MRange.Value = "='175'!H1";
          MRange = workSheet.get_Range("H44");
          MRange.Value = "='175'!I1";
          MRange = workSheet.get_Range("I44");
          MRange.Value = "='175'!J1";
          MRange = workSheet.get_Range("J44");
          MRange.Value = "='175'!K1";
          MRange = workSheet.get_Range("K44");
          MRange.Value = "='175'!L1";
          MRange = workSheet.get_Range("L44");
          MRange.Value = "='175'!M1";

          MRange = workSheet.get_Range("A45");
          MRange.Value = "117";
          MRange = workSheet.get_Range("B45");
          MRange.Value = "Albina Pump Station";
          MRange = workSheet.get_Range("C45");
          MRange.Value = "5.14765576223667%";
          MRange = workSheet.get_Range("F45");
          MRange.Value = "='117'!G1";
          MRange = workSheet.get_Range("G45");
          MRange.Value = "='117'!H1";
          MRange = workSheet.get_Range("H45");
          MRange.Value = "='117'!I1";
          MRange = workSheet.get_Range("I45");
          MRange.Value = "='117'!J1";
          MRange = workSheet.get_Range("J45");
          MRange.Value = "='117'!K1";
          MRange = workSheet.get_Range("K45");
          MRange.Value = "='117'!L1";
          MRange = workSheet.get_Range("L45");
          MRange.Value = "='117'!M1";

          MRange = workSheet.get_Range("A46");
          MRange.Value = "173";
          MRange = workSheet.get_Range("B46");
          MRange.Value = "METRO Learning Center";
          MRange = workSheet.get_Range("C46");
          MRange.Value = "4.70941030475328%";
          MRange = workSheet.get_Range("F46");
          MRange.Value = "='173'!G1";
          MRange = workSheet.get_Range("G46");
          MRange.Value = "='173'!H1";
          MRange = workSheet.get_Range("H46");
          MRange.Value = "='173'!I1";
          MRange = workSheet.get_Range("I46");
          MRange.Value = "='173'!J1";
          MRange = workSheet.get_Range("J46");
          MRange.Value = "='173'!K1";
          MRange = workSheet.get_Range("K46");
          MRange.Value = "='173'!L1";
          MRange = workSheet.get_Range("L46");
          MRange.Value = "='173'!M1";

          MRange = workSheet.get_Range("A47");
          MRange.Value = "192";
          MRange = workSheet.get_Range("B47");
          MRange.Value = "Children's Museum";
          MRange = workSheet.get_Range("C47");
          MRange.Value = "4.2058285030458%";
          MRange = workSheet.get_Range("F47");
          MRange.Value = "='192'!G1";
          MRange = workSheet.get_Range("G47");
          MRange.Value = "='192'!H1";
          MRange = workSheet.get_Range("H47");
          MRange.Value = "='192'!I1";
          MRange = workSheet.get_Range("I47");
          MRange.Value = "='192'!J1";
          MRange = workSheet.get_Range("J47");
          MRange.Value = "='192'!K1";
          MRange = workSheet.get_Range("K47");
          MRange.Value = "='192'!L1";
          MRange = workSheet.get_Range("L47");
          MRange.Value = "='192'!M1";

          MRange = workSheet.get_Range("A48");
          MRange.Value = "213";
          MRange = workSheet.get_Range("B48");
          MRange.Value = "Madison School";
          MRange = workSheet.get_Range("C48");
          MRange.Value = "4.03283199396343%";
          MRange = workSheet.get_Range("F48");
          MRange.Value = "='213'!G1";
          MRange = workSheet.get_Range("G48");
          MRange.Value = "='213'!H1";
          MRange = workSheet.get_Range("H48");
          MRange.Value = "='213'!I1";
          MRange = workSheet.get_Range("I48");
          MRange.Value = "='213'!J1";
          MRange = workSheet.get_Range("J48");
          MRange.Value = "='213'!K1";
          MRange = workSheet.get_Range("K48");
          MRange.Value = "='213'!L1";
          MRange = workSheet.get_Range("L48");
          MRange.Value = "='213'!M1";

          Excel.FormatCondition format1 = (Excel.FormatCondition)(workSheet.get_Range("F36:F48",
      Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreaterEqual,
      workSheet.get_Range("$D$15", Type.Missing), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

          format1.Font.Bold = true;
          format1.Font.Color = 0x000000FF;

          Excel.FormatCondition format2 = (Excel.FormatCondition)(workSheet.get_Range("G36:G48",
      Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Excel.XlFormatConditionOperator.xlGreaterEqual,
        workSheet.get_Range("$E$15", Type.Missing), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

          format2.Font.Bold = true;
          format2.Font.Color = 0x000000FF;

          Excel.FormatCondition format3 = (Excel.FormatCondition)(workSheet.get_Range("H36:H48",
      Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Excel.XlFormatConditionOperator.xlGreaterEqual,
        workSheet.get_Range("$F$15", Type.Missing), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

          format3.Font.Bold = true;
          format3.Font.Color = 0x000000FF;

          Excel.FormatCondition format4 = (Excel.FormatCondition)(workSheet.get_Range("I36:I48",
      Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Excel.XlFormatConditionOperator.xlGreaterEqual,
        workSheet.get_Range("$G$15", Type.Missing), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

          format4.Font.Bold = true;
          format4.Font.Color = 0x000000FF;

          Excel.FormatCondition format5 = (Excel.FormatCondition)(workSheet.get_Range("J36:J48",
      Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Excel.XlFormatConditionOperator.xlGreaterEqual,
        workSheet.get_Range("$H$15", Type.Missing), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

          format5.Font.Bold = true;
          format5.Font.Color = 0x000000FF;

          Excel.FormatCondition format6 = (Excel.FormatCondition)(workSheet.get_Range("K36:K48",
      Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Excel.XlFormatConditionOperator.xlGreaterEqual,
        workSheet.get_Range("$I$15", Type.Missing), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

          format6.Font.Bold = true;
          format6.Font.Color = 0x000000FF;

          Excel.FormatCondition format7 = (Excel.FormatCondition)(workSheet.get_Range("L36:L48",
      Type.Missing).FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Excel.XlFormatConditionOperator.xlGreaterEqual,
        workSheet.get_Range("$J$15", Type.Missing), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

          format7.Font.Bold = true;
          format7.Font.Color = 0x000000FF;


          MRange = workSheet.get_Range("B49");
          MRange.Value = "Willamette Average";
          MRange = workSheet.get_Range("C49");
          MRange.Value = "=SUM(C36:C48)";
          MRange = workSheet.get_Range("F49");
          MRange.Value = "=SUMPRODUCT($C36:$C48,F36:F48)/$C$49";
          MRange = workSheet.get_Range("G49");
          MRange.Value = "=SUMPRODUCT($C36:$C48,G36:G48)/$C$49";
          MRange = workSheet.get_Range("H49");
          MRange.Value = "=SUMPRODUCT($C36:$C48,H36:H48)/$C$49";
          MRange = workSheet.get_Range("I49");
          MRange.Value = "=SUMPRODUCT($C36:$C48,I36:I48)/$C$49";
          MRange = workSheet.get_Range("J49");
          MRange.Value = "=SUMPRODUCT($C36:$C48,J36:J48)/$C$49";
          MRange = workSheet.get_Range("K49");
          MRange.Value = "=SUMPRODUCT($C36:$C48,K36:K48)/$C$49";
          MRange = workSheet.get_Range("L49");
          MRange.Value = "=SUMPRODUCT($C36:$C48,L36:L48)/$C$49";
        }

        // check fielpath
        if (ExcelFilePath != null && ExcelFilePath != "")
        {
          try
          {
            wb.Save();
            //MessageBox.Show("Excel file saved!");
          }
          catch (Exception ex)
          {
            throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                + ex.Message);
          }
          finally
          {
            excelApp.Quit();
          }
        }
        else    // no filepath is given
        {
          excelApp.Visible = true;
        }
      }
      catch (Exception ex)
      {
        throw new Exception("ExportToExcel: \n" + ex.Message);
      }
      finally
      {
        excelApp.Quit();
      }
    }


    // Export DataTable into an excel file with field names in the header line
    // - Save excel file without ever making it visible if filepath is given
    // - Don't save excel file, just make it visible if no filepath is given
    public static void ExportToExcel(this DataTable Tbl, string shortDate, DateTime theStart, int x, string ExcelFilePath = null, int WorkSheetName = 0)
    {
      try
      {
        int FileNew = 0;

        if (Tbl == null || Tbl.Columns.Count == 0)
          throw new Exception("ExportToExcel: Null or empty input table!\n");

        // load excel, and create a new workbook
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook wb = null;
        Excel._Worksheet workSheet;

        if (WorkSheetName == 1)
        {
          wb = excelApp.Workbooks.Add(Type.Missing);
          workSheet = excelApp.ActiveSheet;
          workSheet.Name = "All Gauges";
          workSheet.get_Range("AA1").Value = "Storm Period, hours";
          workSheet.get_Range("AB1").Value = "1";
          workSheet.get_Range("AC1").Value = "2";
          workSheet.get_Range("AD1").Value = "3";
          workSheet.get_Range("AE1").Value = "6";
          workSheet.get_Range("AF1").Value = "12";
          workSheet.get_Range("AG1").Value = "24";
          workSheet.get_Range("AH1").Value = "48";
          workSheet.get_Range("AI1").Value = "72";

          workSheet.get_Range("A2").Value = "Hours";
          workSheet.get_Range("D2").Value = "1";
          workSheet.get_Range("E2").Value = "2";
          workSheet.get_Range("F2").Value = "3";
          workSheet.get_Range("G2").Value = "6";
          workSheet.get_Range("H2").Value = "12";
          workSheet.get_Range("I2").Value = "24";
          workSheet.get_Range("J2").Value = "48";
          workSheet.get_Range("A3").Value = "ASFO Citywide 4-Per-Winter";
          workSheet.get_Range("D3").Value = "0.24";
          workSheet.get_Range("E3").Value = "0.34";
          workSheet.get_Range("F3").Value = "0.44";
          workSheet.get_Range("G3").Value = "0.65";
          workSheet.get_Range("H3").Value = "0.89";
          workSheet.get_Range("I3").Value = "1.19";
          workSheet.get_Range("J3").Value = "1.53";
          workSheet.get_Range("A4").Value = "3-per-Winter Storm";
          workSheet.get_Range("D4").Value = "0.27";
          workSheet.get_Range("E4").Value = "0.38";
          workSheet.get_Range("F4").Value = "0.49";
          workSheet.get_Range("G4").Value = "0.72";
          workSheet.get_Range("H4").Value = "1.01";
          workSheet.get_Range("I4").Value = "1.35";
          workSheet.get_Range("J4").Value = "1.74";
          workSheet.get_Range("A5").Value = "2-per-Winter Storm";
          workSheet.get_Range("D5").Value = "0.30";
          workSheet.get_Range("E5").Value = "0.43";
          workSheet.get_Range("F5").Value = "0.55";
          workSheet.get_Range("G5").Value = "0.81";
          workSheet.get_Range("H5").Value = "1.17";
          workSheet.get_Range("I5").Value = "1.59";
          workSheet.get_Range("J5").Value = "2.07";
          workSheet.get_Range("A6").Value = "1-per-Winter Storm";
          workSheet.get_Range("D6").Value = "0.35";
          workSheet.get_Range("E6").Value = "0.50";
          workSheet.get_Range("F6").Value = "0.65";
          workSheet.get_Range("G6").Value = "0.97";
          workSheet.get_Range("H6").Value = "1.43";
          workSheet.get_Range("I6").Value = "1.93";
          workSheet.get_Range("J6").Value = "2.55";
          workSheet.get_Range("A7").Value = "ASFO 5-Year Winter";
          workSheet.get_Range("D7").Value = "0.43";
          workSheet.get_Range("E7").Value = "0.62";
          workSheet.get_Range("F7").Value = "0.80";
          workSheet.get_Range("G7").Value = "1.21";
          workSheet.get_Range("H7").Value = "1.81";
          workSheet.get_Range("I7").Value = "2.51";
          workSheet.get_Range("J7").Value = "3.26";
          workSheet.get_Range("A8").Value = "BES Sewer Design Manual - Design Storms";
          workSheet.get_Range("A9").Value = "BES 2-Year Storm";
          workSheet.get_Range("D9").Value = "0.46";
          workSheet.get_Range("E9").Value = "0.64";
          workSheet.get_Range("F9").Value = "0.80";
          workSheet.get_Range("G9").Value = "1.19";
          workSheet.get_Range("H9").Value = "1.78";
          workSheet.get_Range("I9").Value = "2.40";
          workSheet.get_Range("A10").Value = "BES 5-Year Storm";
          workSheet.get_Range("D10").Value = "0.59";
          workSheet.get_Range("E10").Value = "0.80";
          workSheet.get_Range("F10").Value = "0.99";
          workSheet.get_Range("G10").Value = "1.49";
          workSheet.get_Range("H10").Value = "2.18";
          workSheet.get_Range("I10").Value = "2.93";
          workSheet.get_Range("A11").Value = "BES 10-Year Storm";
          workSheet.get_Range("D11").Value = "0.68";
          workSheet.get_Range("E11").Value = "0.92";
          workSheet.get_Range("F11").Value = "1.15";
          workSheet.get_Range("G11").Value = "1.68";
          workSheet.get_Range("H11").Value = "2.45";
          workSheet.get_Range("I11").Value = "3.34";
          workSheet.get_Range("A12").Value = "BES 25-Year Storm";
          workSheet.get_Range("D12").Value = "0.79";
          workSheet.get_Range("E12").Value = "1.06";
          workSheet.get_Range("F12").Value = "1.30";
          workSheet.get_Range("G12").Value = "1.91";
          workSheet.get_Range("H12").Value = "2.81";
          workSheet.get_Range("I12").Value = "3.77";
          workSheet.get_Range("A13").Value = "BES 50-Year Storm";
          workSheet.get_Range("D13").Value = "0.90";
          workSheet.get_Range("E13").Value = "1.18";
          workSheet.get_Range("F13").Value = "1.43";
          workSheet.get_Range("G13").Value = "2.13";
          workSheet.get_Range("H13").Value = "3.14";
          workSheet.get_Range("I13").Value = "4.20";
          workSheet.get_Range("A14").Value = "BES 100-Year Storm";
          workSheet.get_Range("D14").Value = "0.99";
          workSheet.get_Range("E14").Value = "1.30";
          workSheet.get_Range("F14").Value = "1.59";
          workSheet.get_Range("G14").Value = "2.34";
          workSheet.get_Range("H14").Value = "3.42";
          workSheet.get_Range("I14").Value = "4.61";
          workSheet.get_Range("A15").Value = "3-Year Summer Storm";
          workSheet.get_Range("D15").Value = "0.40";
          workSheet.get_Range("E15").Value = "0.52";
          workSheet.get_Range("F15").Value = "0.60";
          workSheet.get_Range("G15").Value = "0.85";
          workSheet.get_Range("H15").Value = "1.10";
          workSheet.get_Range("I15").Value = "1.41";
          workSheet.get_Range("A16").Value = "10-Year Summer Storm";
          workSheet.get_Range("D16").Value = "0.51";
          workSheet.get_Range("E16").Value = "0.70";
          workSheet.get_Range("F16").Value = "0.85";
          workSheet.get_Range("G16").Value = "1.25";
          workSheet.get_Range("H16").Value = "1.68";
          workSheet.get_Range("I16").Value = "2.06";
          wb.SaveAs(ExcelFilePath);
          excelApp.Quit();
          excelApp = new Excel.Application();
        }

        try
        {
          wb = excelApp.Workbooks.Open(ExcelFilePath);
          workSheet = excelApp.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
          FileNew = 0;
        }
        catch (Exception ex)
        {
          FileNew = 1;
          excelApp.Workbooks.Add();
          workSheet = excelApp.ActiveSheet;
        }

        // single worksheet
        workSheet.Name = WorkSheetName.ToString() + "_" + shortDate.Replace('/', '_');

        // column headings
        for (int i = 0; i < Tbl.Columns.Count; i++)
        {
          workSheet.Cells[1, (i + 1)] = Tbl.Columns[i].ColumnName;
        }

        // rows
        for (int i = 0; i < Tbl.Rows.Count; i++)
        {
          // to do: format datetime values before printing
          for (int j = 0; j < Tbl.Columns.Count; j++)
          {
            workSheet.Cells[(i + 2), (j + 1)] = Tbl.Rows[i][j];
          }
        }

        //Create chart
        Excel.Range chartRange;

        Excel.ChartObjects xlCharts = (Excel.ChartObjects)workSheet.ChartObjects(Type.Missing);
        Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
        Excel.Chart chartPage = myChart.Chart;

        chartRange = workSheet.get_Range("D:D, E:E", Type.Missing);//"D2:D120, E2:E120"

        chartPage.SetSourceData(chartRange, Type.Missing);
        chartPage.ChartType = Excel.XlChartType.xlXYScatterLines;
        chartPage.Location(Excel.XlChartLocation.xlLocationAsNewSheet, WorkSheetName.ToString() + "_Chart");

        Excel.Range sumRange;

        if (workSheet.get_Range("E13").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("F13");
          sumRange.Formula = "=IF(ISBLANK(E2),TRUE,SUM(E2:E13))";
          sumRange.AutoFill(workSheet.get_Range("F13", "F" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("F1");
          sumRange.Formula = "=MAX(F13:F" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("E25").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("G25");
          sumRange.Formula = "=IF(ISBLANK(E2),TRUE,SUM(E2:E25))";
          sumRange.AutoFill(workSheet.get_Range("G25", "G" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("G1");
          sumRange.Formula = "=MAX(G25:F" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("E37").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("H37");
          sumRange.Formula = "=IF(ISBLANK(E2),TRUE,SUM(E2:E37))";
          sumRange.AutoFill(workSheet.get_Range("H37", "H" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("H1");
          sumRange.Formula = "=MAX(H37:H" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("E73").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("I73");
          sumRange.Formula = "=IF(ISBLANK(E2),TRUE,SUM(E2:E73))";
          sumRange.AutoFill(workSheet.get_Range("I73", "I" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("I1");
          sumRange.Formula = "=MAX(I73:I" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("E145").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("J145");
          sumRange.Formula = "=IF(ISBLANK(E2),TRUE,SUM(E2:E145))";
          sumRange.AutoFill(workSheet.get_Range("J145", "J" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("J1");
          sumRange.Formula = "=MAX(J145:J" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("E289").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("K289");
          sumRange.Formula = "=IF(ISBLANK(E2),TRUE,SUM(E2:E289))";
          sumRange.AutoFill(workSheet.get_Range("K289", "K" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("K1");
          sumRange.Formula = "=MAX(K289:K" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("E577").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("L577");
          sumRange.Formula = "=IF(ISBLANK(E2),TRUE,SUM(E2:E577))";
          sumRange.AutoFill(workSheet.get_Range("L577", "L" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("L1");
          sumRange.Formula = "=MAX(L577:L" + (Tbl.Rows.Count + 1).ToString();
        }

        if (workSheet.get_Range("E865").Cells.Value > -1)
        {
          sumRange = workSheet.get_Range("M865");
          sumRange.Formula = "=IF(ISBLANK(E2),TRUE,SUM(E2:E865))";
          sumRange.AutoFill(workSheet.get_Range("M865", "M" + (Tbl.Rows.Count + 1).ToString()), Excel.XlAutoFillType.xlFillSeries);

          sumRange = workSheet.get_Range("M1");
          sumRange.Formula = "=MAX(M865:M" + (Tbl.Rows.Count + 1).ToString();
        }

        // Create summary chart
        //Create chart
        //Excel.ChartObjects xlCharts = (Excel.ChartObjects)workSheet.ChartObjects(Type.Missing);
        Excel.ChartObject myChart2 = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
        Excel.Chart chartPage2 = myChart2.Chart;

        //chartRange = workSheet.get_Range("F1:M1, 'All Gauges'!D2:J2", Type.Missing);//"D2:D120, E2:E120"
        if (theStart.Month >= 5 && theStart.Month < 11)
        {
          chartRange = ((Excel.Worksheet)wb.Sheets["All Gauges"]).get_Range("D15:I15, AB" + (WorkSheetName + 1).ToString() + ":AI" + (WorkSheetName + 1).ToString());
        }
        else
        {
          chartRange = ((Excel.Worksheet)wb.Sheets["All Gauges"]).get_Range("D3:J3, AB" + (WorkSheetName + 1).ToString() + ":AI" + (WorkSheetName + 1).ToString());
        }

        chartPage2.SetSourceData(chartRange, Type.Missing);

        if (theStart.Month >= 5 && theStart.Month < 11)
        {
          ((Excel.Series)chartPage2.SeriesCollection(1)).Name = "3-year Summer Storm";
          //((Excel.Series)chartPage2.SeriesCollection(1)). = "3-year Summer Storm";
          ((Excel.Series)chartPage2.SeriesCollection(2)).Name = "Event Rainfall";
        }
        else
        {
          ((Excel.Series)chartPage2.SeriesCollection(1)).Name = "ASFO Citywide 4-Per-Winter";
          ((Excel.Series)chartPage2.SeriesCollection(2)).Name = "Event Rainfall";
        }

        chartPage2.ChartType = Excel.XlChartType.xlXYScatterLines;
        chartPage2.Location(Excel.XlChartLocation.xlLocationAsNewSheet, WorkSheetName.ToString() + "_SummaryChart");



        // check fielpath
        if (ExcelFilePath != null && ExcelFilePath.Length > 0)
        {
          try
          {
            wb.Save();
          }
          catch (Exception ex)
          {
            throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                + ex.Message);
          }
          finally
          {
            excelApp.Quit();
          }
        }
        else    // no filepath is given
        {
          excelApp.Visible = true;
        }
      }
      catch (Exception ex)
      {
        throw new Exception("ExportToExcel: \n" + ex.Message);
      }
    }
  }
}
