using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MathNet.Numerics.LinearAlgebra;
using OfficeOpenXml;

namespace Helpers
{
	public static class Helpers
	{
		public static Tuple<Matrix<double>, Matrix<double>, Matrix<double>, Matrix<double>> readData (string filename)
		{
			var fileinfo = new FileInfo (filename);

			if (!fileinfo.Exists)
			{
				throw new FileNotFoundException();
			}

			using (var package = new ExcelPackage (fileinfo)) {
				var wb = package.Workbook;

				using (var ws = wb?.Worksheets ["CrossHoldings"]) {
					if (ws == null)
						return null;

					int length = ws.Dimension.End.Column - 1,
					matrixSize = length * length,
					vectorSize = length,
					startRow = ws.Dimension.Start.Row + 1,
					startColumn = ws.Dimension.Start.Column + 1,
					endRow = ws.Dimension.End.Row,
					endColumn = ws.Dimension.End.Column;

					List<double> vals = ws.Cells [startRow, startColumn, endRow, endColumn]
						.Where (x => x != null)
						.Select (x => x.Value)
						.OfType<double> ()
						.ToList ();

					Matrix<double> shareholdings, outside, remainder, dividends;
					shareholdings = Matrix<double>.Build.DenseOfColumnMajor (length, length, vals.Take (matrixSize));
					outside = Matrix<double>.Build.DenseOfDiagonalArray (vals.Skip (matrixSize).Take (vectorSize).ToArray ());
					remainder = Matrix<double>.Build.DenseOfDiagonalArray (vals.Skip (matrixSize + vectorSize).Take (vectorSize).ToArray ());
					dividends = Matrix<double>.Build.DenseOfColumnMajor (length, 1, vals.Skip (matrixSize + 2 * vectorSize).Take (vectorSize));

					return System.Tuple.Create (shareholdings.Transpose (), outside, remainder, dividends);
				}
			}
		}

		private static void writeMatrix (string sheetname, ExcelWorkbook wb, Matrix<double> src, IEnumerable<string> rows = null, IEnumerable<string> cols = null)
		{
			var ws = wb.Worksheets [sheetname] ?? wb.Worksheets.Add (sheetname);
			var rowIndex = -1;
			var colIndex = -1;

			foreach (var row in src.EnumerateRowsIndexed()) {
				rowIndex = row.Item1 + 2;
				foreach (var val in row.Item2.EnumerateIndexed()) {
					colIndex = val.Item1 + 2;
					double num = val.Item2;
					num = num > 10e-8 ? num : 0.0;

					ws.Cells [rowIndex, colIndex].Value = num;
				}
			}

			if (cols != null) {
				colIndex = 2;
				foreach (var col in cols) {
					ws.Cells [1, colIndex++].Value = col;
				}
			}

			if (rows != null) {
				rowIndex = 2;
				foreach (var row in rows) {
					ws.Cells [rowIndex++, 1].Value = row;
				}
			}
		}

		public static void writeData (string filename, Matrix<double> ownership, Matrix<double> dividends, Matrix<double> exit)
		{
			var fileinfo = new FileInfo (filename);

			using (var package = new ExcelPackage (fileinfo)) {
				var wb = package.Workbook;

				using (var ws = wb?.Worksheets ["CrossHoldings"]) {
					if (ws == null)
						return;
					
					var range = ws.Dimension.End.Column;
					List<string> companies = ws.Cells [2, 1, range, 1]
						.Select (x => x.Value)
						.OfType<string> ()
						.ToList ();
					
					writeMatrix ("OwnershipTable", wb, ownership, Enumerable.Repeat (companies, 2).SelectMany (x => x), companies);
					writeMatrix ("DynamicDividendFlow", wb, dividends, Enumerable.Repeat (companies, 3).SelectMany (x => x), Enumerable.Range (1, dividends.ColumnCount).Select (x => x.ToString ()));
					writeMatrix ("ExitTable", wb, exit, Enumerable.Repeat (companies, 2).SelectMany (x => x), companies);

					package.Save ();
				}
			}
		}
	}
}
