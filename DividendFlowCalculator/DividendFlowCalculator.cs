using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MathNet.Numerics.LinearAlgebra;
using OfficeOpenXml;
using NLog;

namespace DividendFlowCalculator
{
	public class DividendData
	{
		public Matrix<double> shareholdings { get; set; }
		public Matrix<double> outside { get; set; }
		public Matrix<double> remainder { get; set; }
		public Matrix<double> dividends { get; set; }
		public List<string> companies { get; set; } = new List<string> ();
		public Matrix<double> initialDividends {
			get {
				int n = shareholdings.ColumnCount;
				Matrix<double> zeroVector = Matrix<double>.Build.Dense (1, n, 0.0);

				var d0 = dividends.Transpose ().Append (zeroVector).Append (zeroVector).Transpose ();
				return d0;
			}
		}
		public Matrix<double> flowMatrix {
			get {
				int n = shareholdings.ColumnCount;
				Matrix<double> zeroMatrix = Matrix<double>.Build.Dense (n, n, 0.0);
				Matrix<double> id = Matrix<double>.Build.DenseIdentity (n);

				var d01 = shareholdings.Append (zeroMatrix).Append (zeroMatrix);
				var d02 = outside.Append (id).Append (zeroMatrix);
				var d03 = remainder.Append (zeroMatrix).Append (id);
				var f = d01.Transpose ().Append (d02.Transpose ()).Append (d03.Transpose ()).Transpose ();
				return f;
			}
		}
		private const double error_tolerance = 10e-6;

		public bool LoadFromFile (string filename)
		{
			var fileinfo = new FileInfo (filename);

			if (!fileinfo.Exists) {
				throw new FileNotFoundException ();
			}

			using (var package = new ExcelPackage (fileinfo)) {
				var wb = package.Workbook;

				using (var ws = wb?.Worksheets ["CrossHoldings"]) {
					if (ws == null)
						return false;

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

					var range = ws.Dimension.End.Column;
					companies = ws.Cells [2, 1, range, 1]
						.Select (x => x.Value)
						.OfType<string> ()
						.ToList ();

					shareholdings = Matrix<double>.Build.DenseOfColumnMajor (length, length, vals.Take (matrixSize)).Transpose ();
					outside = Matrix<double>.Build.DenseOfDiagonalArray (vals.Skip (matrixSize).Take (vectorSize).ToArray ());
					remainder = Matrix<double>.Build.DenseOfDiagonalArray (vals.Skip (matrixSize + vectorSize).Take (vectorSize).ToArray ());
					dividends = Matrix<double>.Build.DenseOfColumnMajor (length, 1, vals.Skip (matrixSize + 2 * vectorSize).Take (vectorSize));

					return true;
				}
			}
		}

		public Tuple<List<string>, List<string>, Matrix<double>> dynamicDividend ()
		{
			Matrix<double> current_dividends = this.initialDividends;
			Matrix<double> later_dividends = this.flowMatrix * this.initialDividends;
			Matrix<double> div = current_dividends.Append (later_dividends);

			while ((current_dividends - later_dividends).L2Norm () > error_tolerance) {
				current_dividends = later_dividends;
				later_dividends = this.flowMatrix.Multiply (current_dividends);
				div = div.Append (later_dividends);
			}

			return System.Tuple.Create (Enumerable.Repeat (companies, 3).SelectMany (x => x).ToList(), Enumerable.Range(1, div.ColumnCount).Select(n => n.ToString()).ToList(), div);
		}

		public Tuple<List<string>, List<string>, Matrix<double>> ownership ()
		{
			int n = this.flowMatrix.ColumnCount / 3;

			Matrix<double> currentF = this.flowMatrix;
			Matrix<double> nextF = this.flowMatrix.Multiply (currentF);

			while ((nextF - currentF).ColumnNorms (2.0).AbsoluteMaximum () > error_tolerance) {
				currentF = nextF;
				nextF = this.flowMatrix.Multiply (nextF);
			}

			Matrix<double> s = nextF.SubMatrix (n, 3 * n - n, 0, n);

			return System.Tuple.Create (Enumerable.Repeat (companies, 2).SelectMany (x => x).ToList(), companies, s);
		}

		public Tuple<List<string>, List<string>, Matrix<double>> penultimateExit ()
		{
			int n = shareholdings.ColumnCount;
			var h = outside.Multiply (shareholdings).Transpose ().Append (remainder.Multiply (shareholdings).Transpose ()).Transpose ();
			var k = outside.Append (remainder).Transpose ();
			var dtable = this.dynamicDividend ();
			var div = dtable.Item3;

			var e0 = Matrix<double>.Build.Dense (2 * n, n, 0.0);

			for (int i = 0; i < 2 * n; ++i) {
				for (int j = 0; j < n; ++j) {
					e0 [i, j] = k [i, j] * div [j, 0];
				}
			}

			var e = e0;
			var ecurrent = Matrix<double>.Build.Dense (2 * n, n, 0.0);

			for (int t = 0; t < div.ColumnCount; ++t) {
				for (int i = 0; i < 2 * n; ++i) {
					for (int j = 0; j < n; ++j) {
						ecurrent [i, j] = h [i, j] * div [j, t];
					}
				}
				e += ecurrent;
			}
			return System.Tuple.Create (Enumerable.Repeat (companies, 2).SelectMany (x => x).ToList(), companies, e);
		}

		private bool writeMatrix (ExcelWorkbook wb, string sheetname, Tuple<List<string>, List<string>, Matrix<double>> src)
		{
			var ws = wb.Worksheets [sheetname] ?? wb.Worksheets.Add (sheetname);
			var rowIndex = -1;
			var colIndex = -1;

			foreach (var row in src.Item3.EnumerateRowsIndexed()) {
				rowIndex = row.Item1 + 2;
				foreach (var val in row.Item2.EnumerateIndexed()) {
					colIndex = val.Item1 + 2;
					double num = val.Item2;
					num = num > 10e-8 ? num : 0.0;

					ws.Cells [rowIndex, colIndex].Value = num;
				}
			}

			var rows = src.Item1;
			var cols = src.Item2;

			if (rows != null) {
				rowIndex = 2;
				foreach (var row in rows) {
					ws.Cells [rowIndex++, 1].Value = row;
				}
			}

			if (cols != null) {
				colIndex = 2;
				foreach (var col in cols) {
					ws.Cells [1, colIndex++].Value = col;
				}
			}
			return true;
		}

		public bool writeData (string filename)
		{
			var fileinfo = new FileInfo (filename);

			using (var package = new ExcelPackage (fileinfo)) {
				var wb = package.Workbook;

				bool hasAnyInvalid = false;
				hasAnyInvalid |= !writeMatrix (wb, "Ownership", ownership ());
				hasAnyInvalid |= !writeMatrix (wb, "DynamicDividendFlow", dynamicDividend ());
				hasAnyInvalid |= !writeMatrix (wb, "PenultimateExit", penultimateExit ());

				package.Save ();
				return hasAnyInvalid;
			}
		}
	}
}
