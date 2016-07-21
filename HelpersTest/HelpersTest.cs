using NUnit.Framework;
using System;
using System.IO;
using Helpers;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;
using System.Reflection;

namespace GivenAFile
{
	[TestFixture]
	public class OnLoad
	{
		[Test]
		public void Success ()
		{
			var vals = Helpers.Helpers.readData ("success.xlsx");

			Matrix<double> c, o, r, g;
			int n = 0;
			int length = 3;

			c = vals.Item1;
			n = c.ColumnCount;
			Assert.AreEqual (n, length);

			o = vals.Item2;
			r = vals.Item3;
			g = vals.Item4;
			Matrix<double> shareholdings, outside, remainder, dividends;
			shareholdings = DenseMatrix.OfArray (new double[,] { { 0, 0.1, 0.2 }, { 0, 0, 0.15 }, { 0.12, 0.12, 0 } });
			Assert.AreEqual(c, shareholdings);

			outside = DenseMatrix.Create (length, length, 0.0);
			outside.SetDiagonal (new double[] { 0.11, 0.21, 0.13 });
			Assert.AreEqual(o, outside);

			remainder = DenseMatrix.Create (length, length, 0.0);
			remainder.SetDiagonal (new double[] { 0.77, 0.57, 0.52 });
			Assert.AreEqual(r, remainder);

			dividends = DenseMatrix.OfColumnMajor (length, 1, new double[] {1000, 0, 0});
			Assert.AreEqual(g, dividends);
		}

		[Test]
		public void MissingFile ()
		{
			Assert.Throws<System.IO.FileNotFoundException> (() => Helpers.Helpers.readData ("missing.xlsx"));
		}

		[Test]
		public void InvalidData () {
			var vals = Helpers.Helpers.readData ("invalid.xlsx");
			Assert.IsNull (vals);
		}
	}
}
