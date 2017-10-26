using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using MathNet.Numerics.LinearAlgebra;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using DividendFlowCalculator;

namespace EconToolsExcel
{
    public partial class ThisAddIn
    {
        private DividendData dd = new DividendData();
       
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private bool WriteMatrixToWorksheet(Excel.Worksheet worksheet, Matrix<double> data, List<string> rowNames, List<string> colNames)
        {
            foreach (var item in data.EnumerateIndexed())
            {
                var row = item.Item1 + 2;
                var col = item.Item2 + 2;
                var num = item.Item3;
                var cell = (Excel.Range)worksheet.Cells[row, col];
                worksheet.Cells[row, col].Value = num;
            }

            if (rowNames != null)
            {
                int rowIndex = 2;
                foreach (var row in rowNames)
                {
                    worksheet.Cells[rowIndex++, 1].Value = row;
                }
            }

            if (colNames != null)
            {
                int colIndex = 2;
                foreach (var col in colNames)
                {
                    worksheet.Cells[1, colIndex++].Value = col;
                }
            }
            return true;
        }

        private void ReadData()
        {
            Excel.Range selection = Application.Selection as Excel.Range;

            int rows = selection.Rows.Count;
            int columns = selection.Columns.Count;

            int length = columns;
            int matrixSize = length * length;
            int vectorSize = length;

            List<double> vals = new List<double>();

            for (int rowIndex = 1; rowIndex <= rows; ++rowIndex)
            {
                for (int colIndex = 1; colIndex <= columns; ++colIndex)
                {
                    Excel.Range cell = selection.get_Item(rowIndex, colIndex) as Excel.Range;
                    if (cell.Value2 != null)
                    {
                        vals.Add(cell.Value2);
                    }
                }
            }

            dd.shareholdings = Matrix<double>.Build.DenseOfRowMajor(length, length, vals.Take(matrixSize));
            dd.outside = Matrix<double>.Build.DenseOfDiagonalArray(vals.Skip(matrixSize).Take(vectorSize).ToArray());
            dd.remainder = Matrix<double>.Build.DenseOfDiagonalArray(vals.Skip(matrixSize + vectorSize).Take(vectorSize).ToArray());
            dd.dividends = Matrix<double>.Build.DenseOfColumnMajor(length, 1, vals.Skip(matrixSize + 2 * vectorSize).Take(vectorSize));
        }

        private void ReadCompanies()
        {
            Excel.Range selection = Application.Selection as Excel.Range;
            int columns = selection.Columns.Count;
            int length = columns;

            dd.companies = new List<string>();
            for (int companyIndex = 2; companyIndex <= length + 1; ++companyIndex)
            {
                dd.companies.Add(Application.ActiveSheet.Cells[1, companyIndex].Value);
            }
        }

        public void AddOwnershipWorksheet()
        {
            ReadData();
            ReadCompanies();

            Excel.Worksheet ownershipWorksheet;
            ownershipWorksheet = (Excel.Worksheet)this.Application.Worksheets.Add(After: Application.ActiveWorkbook.Sheets[Application.ActiveWorkbook.Sheets.Count]);
            ownershipWorksheet.Name = "Ownership";

            WriteMatrixToWorksheet(ownershipWorksheet, dd.ownership, Enumerable.Repeat(dd.companies, 2).SelectMany(x => x).ToList(), dd.companies);
        }

        public void AddDividendWorksheet()
        {
            ReadData();
            ReadCompanies();

            Excel.Worksheet dynamicDividendFlow;
            dynamicDividendFlow = (Excel.Worksheet)this.Application.Worksheets.Add(After: Application.ActiveWorkbook.Sheets[Application.ActiveWorkbook.Sheets.Count]);
            dynamicDividendFlow.Name = "DynamicDividendFlow";

            WriteMatrixToWorksheet(dynamicDividendFlow, dd.dynamicDividend, Enumerable.Repeat(dd.companies, 3).SelectMany(x => x).ToList(), Enumerable.Range(1, dd.dynamicDividend.ColumnCount).Select(x => x.ToString()).ToList());
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
