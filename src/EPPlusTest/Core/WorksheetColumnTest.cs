using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlusTest.Core
{
    [TestClass]
    public class WorksheetColumnTest : TestBase
    {
        [ClassInitialize]
        public static void Init(TestContext context)
        {
         //   _pck = OpenPackage("ColumnTests.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
           // SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void ValidateDefaultWidth()
        {
            using(var p = OpenPackage("columnWidthDefault.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("default");
                var expectedWidth = 9.140625D;
                Assert.AreEqual(expectedWidth, ws.DefaultColWidth);

                ws.Column(2).Width = ws.DefaultColWidth;
                SaveAndCleanup(p);
            }
        }        
        [TestMethod, Ignore]
        public void ValidateWidthHeeboLight()
        {
            CreateNormalFontsFiles("Heebo Light");
        }
        [TestMethod, Ignore]
        public void ValidateWidthVerdana()
        {
            CreateNormalFontsFiles("Verdana");
        }
        [TestMethod, Ignore]
        public void ValidateWidthArial()
        {
            CreateNormalFontsFiles("Arial");
        }
        [TestMethod, Ignore]
        public void ValidateWidthCalibri()
        {
            CreateNormalFontsFiles("Calibri");
        }
        [TestMethod, Ignore]
        public void ValidateWidthTimesNewRoman()
        {
            CreateNormalFontsFiles("Times New Roman");
        }        
        private static void CreateNormalFontsFiles(string fontName)
        {
            var fontNameNoSpace = fontName.Replace(" ", "");
            foreach (var size in new int[] { 6, 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42, 44, 48, 72, 96, 128, 256 })
            {
                using (var p = OpenPackage($"ColumnWidth\\columnWidth{fontNameNoSpace}{size}.xlsx", true))
                {
                    var ws = p.Workbook.Worksheets.Add($"{fontNameNoSpace}{size}");
                    p.Workbook.Styles.NamedStyles[0].Style.Font.Name = fontName;
                    p.Workbook.Styles.NamedStyles[0].Style.Font.Size = size;

                    ws.Column(2).Width = ws.DefaultColWidth;
                    SaveAndCleanup(p);
                }
            }
        }
        [TestMethod]
        public void ValidateAutoFitWidthNormalArial28()
        {
            using (var p = OpenPackage($"columnWidthArial28.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add($"arial28");
                p.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Arial";
                p.Workbook.Styles.NamedStyles[0].Style.Font.Size = 28;

                ws.Cells["A1"].Value = "12345678";
                ws.Column(1).AutoFit();

                ws.Column(2).Width = ws.DefaultColWidth;
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ValidateDefaultWidthArial36()
        {
            using (var p = OpenPackage("columnWidthArial36.xlsx", true))
            {   
                var ws = p.Workbook.Worksheets.Add("arial36");
                p.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Arial";
                p.Workbook.Styles.NamedStyles[0].Style.Font.Size = 36;

                ws.Column(2).Width = ws.DefaultColWidth;
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ValidateDefaultWidthArial72()
        {
            using (var p = OpenPackage("columnWidthArial72.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("arial72");
                p.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Arial";
                p.Workbook.Styles.NamedStyles[0].Style.Font.Size = 72;

                ws.Column(2).Width = ws.DefaultColWidth;
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ColumnCheck()
        {
            using (var p = OpenTemplatePackage("s808_2.xlsx"))
            {
                var ws = p.Workbook.Worksheets["overzicht"];

                ws.Calculate();
                ws.ClearFormulas();

                List<ExcelRangeColumn> cols = [.. ws.Columns.Where(c => c.Hidden).OrderByDescending(c => c.StartColumn)];

                List<int> deletedCols = new();

                foreach (ExcelRangeColumn col in cols)
                {
                    ws.DeleteColumn(col.StartColumn);
                    deletedCols.Add(col.StartColumn);
                }

                foreach (var drawing in ws.Drawings)
                {
                    if (drawing.DrawingType == eDrawingType.Chart)
                    {
                        var chartSerie = drawing.As.Chart.Chart.Series;

                        foreach (var serie in chartSerie)
                        {
                            foreach (var col in deletedCols)
                            {
                                DeleteColumnFromSeries(ws, serie, col);
                            }
                        }
                    }
                }
                SaveAndCleanup(p);
            }
        }

        public void DeleteColumnFromSeries(ExcelWorksheet ws, ExcelChartSerie serie, int deletedColumn)
        {
            if (serie.HeaderAddress != null)
            {
                serie.HeaderAddress = ws.Cells[UpdateSerieString(ws, serie.HeaderAddress.Address, deletedColumn)];
            }
            serie.Series = UpdateSerieString(ws, serie.Series, deletedColumn);
            serie.XSeries = UpdateSerieString(ws, serie.XSeries, deletedColumn);
        }

        public string UpdateSerieString(ExcelWorksheet ws, string serieString, int deletedColumn)
        {
            string updatedString = serieString;

            if (!string.IsNullOrEmpty(serieString))
            {
                var newSerieString = DeleteColumnFromAddress(ws, new ExcelAddress(serieString), deletedColumn);

                if (newSerieString != null)
                {
                    updatedString = newSerieString;
                }
            }

            return updatedString;
        }

        public string DeleteColumnFromAddress(ExcelWorksheet ws, ExcelAddressBase address, int deletedColumn)
        {
            if (address != null)
            {
                if (address.Start.Column > deletedColumn)
                {
                    var start = address.Start;
                    var end = address.End;

                    var newAddress = ws.Cells[start.Row, start.Column - 1, end.Row, end.Column - 1];
                    return newAddress.FullAddressAbsolute;
                }
                return address.Address;
            }
            return null;
        }
    }
}
