using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WebApplicationTest
{
    public class TableToExcel
    {
        ExcelPackage excel = new ExcelPackage();
        ExcelWorksheet sheet;
        private int maxRow = 0;
        private Dictionary<string, object> cellsOccupied = new Dictionary<string, object>();

        public TableToExcel()
        {
            sheet = excel.Workbook.Worksheets.Add("sheet1");
            sheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells.Style.WrapText = true;
        }

        public byte[] process(string html)
        {
            MemoryStream stream = null;
            try
            {
                process(html, out stream);
                return stream.ToArray();
            }
            finally
            {
                if (stream != null)
                {
                    try
                    {
                        stream.Close();
                    }
                    catch (IOException e)
                    {
                        throw e;
                    }
                }
            }
        }

        public void process(String html, out MemoryStream output)
        {
            WebBrowser wb = new WebBrowser();
            wb.Navigate("about:blank");
            HtmlDocument doc = wb.Document.OpenNew(true);
            doc.Write(html);
            foreach (HtmlElement table in doc.GetElementsByTagName("table"))
            {
                processTable(table);
            }
            try
            {
                output = new MemoryStream();
                excel.SaveAs(output);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private void processTable(HtmlElement table)
        {
            int rowIndex = 1;
            if (maxRow > 0)
            {
                // blank row
                maxRow += 1;
                rowIndex = maxRow;
            }
            // Interate Table Rows.
            foreach (HtmlElement row in table.GetElementsByTagName("tr"))
            {
                int colIndex = 1;
                // Interate Cols.
                HtmlElementCollection tds = row.GetElementsByTagName("th");
                if (tds.Count <= 0)
                {
                    tds = row.GetElementsByTagName("td");
                }
                foreach (HtmlElement td in tds)
                {
                    // skip occupied cell
                    while (cellsOccupied.ContainsKey(rowIndex + "_" + colIndex))
                    {
                        ++colIndex;
                    }
                    int rowSpan = getSpan(td.OuterHtml, 0);
                    int colSpan = getSpan(td.OuterHtml, 1);
                    sheet.Cells[rowIndex, colIndex].Value = td.InnerText;
                    // col span & row span
                    if (colSpan > 1 && rowSpan > 1)
                    {
                        spanRowAndCol(rowIndex, colIndex, rowSpan, colSpan);
                        colIndex += colSpan;
                    }
                    // col span only
                    else if (colSpan > 1)
                    {
                        spanCol(rowIndex, colIndex, colSpan);
                        colIndex += colSpan;
                    }
                    // row span only
                    else if (rowSpan > 1)
                    {
                        spanRow(rowIndex, colIndex, rowSpan);
                        ++colIndex;
                    }
                    // no span
                    else
                    {
                        ++colIndex;
                    }
                }
                ++rowIndex;
                if (rowIndex > maxRow)
                {
                    maxRow = rowIndex;
                }
            }
        }

        private void spanRow(int rowIndex, int colIndex, int rowSpan)
        {
            sheet.Cells[rowIndex, colIndex, rowIndex + rowSpan - 1, colIndex].Merge = true;
            for (int i = 0; i < rowSpan; i++)
            {
                cellsOccupied.Add((rowIndex + i) + "_" + colIndex, true);
            }
            if (rowIndex + rowSpan - 1 > maxRow)
            {
                maxRow = rowIndex + rowSpan - 1;
            }
        }

        private void spanCol(int rowIndex, int colIndex, int colSpan)
        {
            sheet.Cells[rowIndex, colIndex, rowIndex, colIndex + colSpan - 1].Merge = true;
        }

        private void spanRowAndCol(int rowIndex, int colIndex, int rowSpan, int colSpan)
        {
            sheet.Cells[rowIndex, colIndex, rowIndex + rowSpan - 1, colIndex + colSpan - 1].Merge = true;
            for (int i = 0; i < rowSpan; i++)
            {
                for (int j = 0; j < colSpan; j++)
                {
                    cellsOccupied.Add((rowIndex + i) + "_" + (colIndex + j), true); 
                }
            }
            if (rowIndex + rowSpan - 1 > maxRow)
            {
                maxRow = rowIndex + rowSpan - 1;
            }
        }

        private int getSpan(string html, int spanType = 0)
        {
            string spanTypeText;
            int span;

            if (spanType == 0)
            {
                spanTypeText = "row";
            }
            else
            {
                spanTypeText = "col";
            }
            string equation = Regex.Match(html.ToLower(), spanTypeText + @"span=.*?\d{1,}").ToString();
            if (!Int32.TryParse(Regex.Match(equation, @"\d{1,}").ToString(), out span))
            {
                span = 1;
            }

            return span;
        }
    }
}
