using System;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;

namespace DividendsCalc
{
	public static class DividendsCalc
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

		public static Matrix<double> penultimateTable (Tuple<Matrix<double>, int> dtable, Matrix<double> c, Matrix<double> o, Matrix<double> r)
		{
			int n = c.ColumnCount;

			var h = o.Multiply (c).Transpose ().Append (r.Multiply (c).Transpose ()).Transpose ();
			var k = o.Append (r).Transpose ();
			var div = dtable.Item1;

			var e0 = DenseMatrix.Create (2 * n, n, 0.0);

			for (int i = 0; i < 2 * n; ++i) {
				for (int j = 0; j < n; ++j) {
					e0 [i, j] = k [i, j] * div [j, 0];
				}
			}

			var e = e0;
			var ecurrent = DenseMatrix.Create (2 * n, n, 0.0);

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
