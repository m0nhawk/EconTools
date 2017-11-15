using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using MathNet.Numerics.LinearAlgebra;
using DividendFlowCalculator;
using System.Windows.Forms;

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

        public void ReadData()
        {
            Excel.Range selection = Application.Selection;

            int start_row, end_row;
            int start_col, end_col;
            int length;
            int matrix_size;
            int vector_size;

            var test = (selection.get_Item(1, 1) as Excel.Range).Cells.Value2;

            dd.companies = new List<string>();
            if (test is double)
            {
                start_row = 1;
                end_row = selection.Rows.Count;

                start_col = 1;
                end_col = selection.Columns.Count;
                
                for (int companyIndex = 2; companyIndex <= end_col + 1; ++companyIndex)
                {
                    dd.companies.Add(Application.ActiveSheet.Cells[1, companyIndex].Value);
                }
            }
            else
            {
                start_row = 2;
                end_row = selection.Rows.Count;

                start_col = 2;
                end_col = selection.Columns.Count;
                
                for (int companyIndex = start_col; companyIndex <= end_col; ++companyIndex)
                {
                    dd.companies.Add(selection[start_row - 1, companyIndex].Value);
                }
            }

            length = end_col - start_col + 1;
            matrix_size = length * length;
            vector_size = length;
            
            List<double> sharehodling = new List<double>();
            List<double> rest = new List<double>();

            for (int rowIndex = start_row; rowIndex <= end_col; ++rowIndex)
            {
                for (int colIndex = start_col; colIndex <= end_col; ++colIndex)
                {
                    Excel.Range cell = selection[rowIndex, colIndex];
                    sharehodling.Add(cell.Value2);
                }
            }

            for (int rowIndex = end_row - 2; rowIndex <= end_row; ++rowIndex)
            {
                for (int colIndex = start_col; colIndex <= end_col; ++colIndex)
                {
                    Excel.Range cell = selection[rowIndex, colIndex];
                    if (cell.Value2 != null)
                    {
                        rest.Add(cell.Value2);
                    }
                    else
                    {
                        rest.Add(0.0);
                    }
                }
            }

            dd.shareholdings = Matrix<double>.Build.DenseOfRowMajor(length, length, sharehodling);
            dd.outside = Matrix<double>.Build.DenseOfDiagonalArray(rest.Take(vector_size).ToArray());
            dd.remainder = Matrix<double>.Build.DenseOfDiagonalArray(rest.Skip(vector_size).Take(vector_size).ToArray());
            dd.dividends = Matrix<double>.Build.DenseOfColumnMajor(length, 1, rest.Skip(2 * vector_size).Take(vector_size));
        }

        public void AddOwnershipWorksheet()
        {
            ReadData();

            Excel.Worksheet ownershipWorksheet;
            ownershipWorksheet = (Excel.Worksheet)this.Application.Worksheets.Add(After: Application.ActiveWorkbook.Sheets[Application.ActiveWorkbook.Sheets.Count]);
            ownershipWorksheet.Name = "Ownership";

            WriteMatrixToWorksheet(ownershipWorksheet, dd.ownership, Enumerable.Repeat(dd.companies, 2).SelectMany(x => x).ToList(), dd.companies);
        }

        public void AddDividendWorksheet()
        {
            ReadData();

            Excel.Worksheet dynamicDividendFlow;
            dynamicDividendFlow = (Excel.Worksheet)this.Application.Worksheets.Add(After: Application.ActiveWorkbook.Sheets[Application.ActiveWorkbook.Sheets.Count]);
            dynamicDividendFlow.Name = "DynamicDividendFlow";

            WriteMatrixToWorksheet(dynamicDividendFlow, dd.dynamicDividend / 100, Enumerable.Repeat(dd.companies, 3).SelectMany(x => x).ToList(), Enumerable.Range(1, dd.dynamicDividend.ColumnCount).Select(x => x.ToString()).ToList());
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
