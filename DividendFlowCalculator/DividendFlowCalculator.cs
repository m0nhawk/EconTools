using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MathNet.Numerics.LinearAlgebra;
using OfficeOpenXml;

namespace DividendFlowCalculator
{
	public static class DividendFlowCalculator
	{
		const double error_tolerance = 10e-6;

		public static Tuple<Matrix<double>, int> dynamicDividendTable (Matrix<double> d0, Matrix<double> f)
		{
			Matrix<double> current_dividends = d0;
			Matrix<double> later_dividends = f * d0;
			Matrix<double> div = current_dividends.Append (later_dividends);

			var T = 2;

			while ((current_dividends - later_dividends).L2Norm () > error_tolerance) {
				current_dividends = later_dividends;
				later_dividends = f.Multiply (current_dividends);
				div = div.Append (later_dividends);
				++T;
			}

			return System.Tuple.Create (div, T);
		}

		public static Matrix<double> ownershipTable (Matrix<double> f)
		{
			int n = f.ColumnCount / 3;

			Matrix<double> currentF = f;
			Matrix<double> nextF = f.Multiply (currentF);

			while ((nextF - currentF).ColumnNorms (2.0).AbsoluteMaximum () > error_tolerance) {
				currentF = nextF;
				nextF = f.Multiply (nextF);
			}

			Matrix<double> s = nextF.SubMatrix (n, 3 * n - n, 0, n);

			return s;
		}

		public static Matrix<double> penultimateExitTable (Tuple<Matrix<double>, int> dtable, Matrix<double> c, Matrix<double> o, Matrix<double> r)
		{
			int n = c.ColumnCount;
			var h = o.Multiply (c).Transpose ().Append (r.Multiply (c).Transpose ()).Transpose ();
			var k = o.Append (r).Transpose ();
			var div = dtable.Item1;

			var e0 = Matrix<double>.Build.Dense (2 * n, n, 0.0);

			for (int i = 0; i < 2 * n; ++i) {
				for (int j = 0; j < n; ++j) {
					e0 [i, j] = k [i, j] * div [j, 0];
				}
			}

			var e = e0;
			var ecurrent = Matrix<double>.Build.Dense (2 * n, n, 0.0);

			for (int t = 0; t < dtable.Item2; ++t) {
				for (int i = 0; i < 2 * n; ++i) {
					for (int j = 0; j < n; ++j) {
						ecurrent [i, j] = h [i, j] * div [j, t];
					}
				}
				e += ecurrent;
			}
			return e;
		}
	}
}

namespace Helpers
{
	public static class Helpers
	{
		public static Tuple<Matrix<double>, Matrix<double>, Matrix<double>, Matrix<double>> readData (string filename)
		{
			var fileinfo = new FileInfo (filename);

			if (!fileinfo.Exists) {
				throw new FileNotFoundException ();
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

		public static Tuple<Matrix<double>, Matrix<double>> ConvertTo (Tuple<Matrix<double>, Matrix<double>, Matrix<double>, Matrix<double>> input)
		{
			int n = input.Item1.ColumnCount;
			Matrix<double> zeroVector = Matrix<double>.Build.Dense (1, n, 0.0);
			Matrix<double> zeroMatrix = Matrix<double>.Build.Dense (n, n, 0.0);
			Matrix<double> id = Matrix<double>.Build.DenseIdentity (n);

			Matrix<double> shareholdings = input.Item1;
			Matrix<double> outside = input.Item2;
			Matrix<double> remainder = input.Item3;
			Matrix<double> dividends = input.Item4;

			var d01 = shareholdings.Append (zeroMatrix).Append (zeroMatrix);
			var d02 = outside.Append (id).Append (zeroMatrix);
			var d03 = remainder.Append (zeroMatrix).Append (id);

			var f = d01.Transpose ().Append (d02.Transpose ()).Append (d03.Transpose ()).Transpose ();
			var d0 = dividends.Transpose ().Append (zeroVector).Append (zeroVector).Transpose ();

			return System.Tuple.Create (f, d0);
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