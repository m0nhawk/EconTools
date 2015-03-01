using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;
using OfficeOpenXml;

namespace Helpers
{
	public static class Helpers
	{
		public static Tuple<Matrix<double>, Matrix<double>, Matrix<double>, Matrix<double>> readData (string filename)
		{
			List<double> vals = new List<double> ();
			int length = 0;
			Matrix<double> shareholdings, outside, remainder, dividends;

			var fileinfo = new FileInfo (filename);

			using (var package = new ExcelPackage (fileinfo)) {
				var wb = package.Workbook;

				if (wb != null) {
					if (wb.Worksheets.Count > 0) {
						using (var ws = wb.Worksheets ["CrossHoldings"]) {
							length = ws.Dimension.End.Column;
							while (ws.Cells [1, length].Value == null) {
								length -= 1;
							}
							length -= 1;

							for (int i = ws.Dimension.Start.Row + 1; i <= ws.Dimension.End.Row; i++) {
								for (int j = ws.Dimension.Start.Column + 1; j <= ws.Dimension.End.Column; j++) {
									if (ws.Cells [i, j].Value != null) {
										double val = ws.Cells [i, j].GetValue<double> ();
										vals.Add (val);
									}
								}
							}
						}
					}
				}
			}

			shareholdings = DenseMatrix.OfColumnMajor (length, length, vals.Take (length * length));

			outside = DenseMatrix.Create (length, length, 0.0);
			outside.SetDiagonal (vals.Skip (length * length).Take (length).ToArray ());

			remainder = DenseMatrix.Create (length, length, 0.0);
			remainder.SetDiagonal (vals.Skip (length * length + length).Take (length).ToArray ());

			dividends = DenseMatrix.OfColumnMajor (length, 1, vals.Skip (length * length + 2 * length).Take (length));

			return System.Tuple.Create (shareholdings.Transpose(), outside, remainder, dividends);
		}

		private static void writeMatrix (string sheetname, ExcelWorkbook wb, Matrix<double> src, IEnumerable<string> rows = null, IEnumerable<string> cols = null)
		{
			var ws = wb.Worksheets [sheetname] ?? wb.Worksheets.Add (sheetname);
			if (cols != null) {
				int i = 2;
				foreach (string col in cols) {
					ws.Cells [1, i++].Value = col;
				}
			}

			var rowIndex = 2;
			var colIndex = 2;

			foreach (System.Tuple<int, Vector<double>> c in src.EnumerateRowsIndexed()) {
				if (rows != null) {
					var enumerable = rows as object[] ?? rows.ToArray ();
					ws.Cells [rowIndex, 1].Value = enumerable.ElementAt (rowIndex - 2);
				}

				colIndex = 2;
				foreach (System.Tuple<int, double> val in c.Item2.EnumerateIndexed()) {
					double num = val.Item2;

					var v = num > 10e-8 ? num : 0.0;

					ws.Cells [rowIndex, colIndex++].Value = v;
				}
				rowIndex += 1;
			}
		}

		public static void writeData (string filename, Matrix<double> ownership, Matrix<double> dividends, Matrix<double> exit)
		{
			var fileinfo = new FileInfo (filename);

			using (var package = new ExcelPackage (fileinfo)) {
				var wb = package.Workbook;

				List<string> companies = new List<string> ();

				var ws = wb.Worksheets ["CrossHoldings"];
				var range = ws.Dimension.End.Column;
				for(int i = 2; i <= range; i++) {
					var company = ws.Cells [i, 1].GetValue<string>();
					if (company != "") {
						companies.Add (company);
					}
				}

				writeMatrix ("OwnershipTable", wb, ownership, Enumerable.Repeat (companies, 2).SelectMany (x => x), companies);
				writeMatrix ("DynamicDividendFlow", wb, dividends, Enumerable.Repeat (companies, 3).SelectMany (x => x), Enumerable.Range (1, dividends.ColumnCount).Select (x => x.ToString ()));
				writeMatrix ("ExitTable", wb, exit, Enumerable.Repeat (companies, 2).SelectMany (x => x), companies);

				Console.Write ("{0} calculated!", filename);

				package.Save ();
			}
		}
	}
}

