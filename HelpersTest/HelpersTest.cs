using NUnit.Framework;
using System;
using MathNet.Numerics.LinearAlgebra;

using Helpers;

namespace GivenAFile
{
	[TestFixture]
	public class OnLoad
	{
		private Matrix<double> shareholdings, outside, remainder, dividends;

		private int length = 3;

		[SetUp]
		public void Init ()
		{
			shareholdings = Matrix<double>.Build.DenseOfArray (new double[,]
				{ { 0, 0.1, 0.2 },
				  { 0, 0, 0.15 },
				  { 0.12, 0.12, 0 } });

			outside = Matrix<double>.Build.DenseOfDiagonalArray (new double[] { 0.11, 0.21, 0.13 });

			remainder = Matrix<double>.Build.DenseOfDiagonalArray (new double[] { 0.77, 0.57, 0.52 });

			dividends = Matrix<double>.Build.DenseOfColumnMajor (length, 1, new double[] { 1000, 0, 0 });
		}

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

			Assert.AreEqual(c, shareholdings);

			Assert.AreEqual(o, outside);

			Assert.AreEqual(r, remainder);

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
