using NUnit.Framework;
using System;
using System.Collections.Generic;
using MathNet.Numerics.LinearAlgebra;

using DividendsCalc;

namespace GivenAData
{
	public class MatrixWithinComparer : IComparer<Matrix<double>>
	{
		private double epsilon = 10e-8;

		public MatrixWithinComparer(double epsilon)
		{
			this.epsilon = epsilon;
		}

		int IComparer<Matrix<double>>.Compare(Matrix<double> a, Matrix<double> b)
		{
			foreach (var v in a.EnumerateIndexed ())
			{
				if (v.Item3 - b.At (v.Item1, v.Item2) > epsilon)
					return -1;
			}
			return 0;
		}
	}

	[TestFixture]
	public class OnCalculation
	{
		private Matrix<double> shareholdings, outside, remainder, dividends;
		private Tuple<Matrix<double>, int> dynamicDividend;
		private Matrix<double> ownership;
		private Matrix<double> penultimateExit;
		private Matrix<double> d0, f;

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

			ownership = Matrix<double>.Build.DenseOfArray (new double[,]
				{ { 0.1129679439, 0.0142647149, 0.0247332029 },
				  { 0.0039530958, 0.2143483761, 0.0329427647 },
				  { 0.016314512, 0.0179459484, 0.1359547246 },
				  { 0.7907756072, 0.099853004, 0.17313242 },
				  { 0.0107298313, 0.5818027351, 0.0894160756 },
				  { 0.0652580481, 0.0717837935, 0.5438188982 }});

			penultimateExit = Matrix<double>.Build.DenseOfArray (new double[,]
				{ { 110, 0.2070696487, 2.7609286707 },
				  { 0, 0, 3.9531478695 },
				  { 16.0209161215, 0.2936624109, 0 },
				  { 770, 1.4494875412, 19.3265006951 },
				  { 0, 0, 10.7299727885 },
				  { 64.0836644861, 1.1746496438, 0 }});
		}

		[Test]
		public void CorrectDynamicDidivend ()
		{
			Matrix<double> zeroVector = Matrix<double>.Build.Dense (1, length, 0.0);
			Matrix<double> zeroMatrix = Matrix<double>.Build.Dense (length, length, 0.0);
			Matrix<double> id = Matrix<double>.Build.DenseIdentity (length);

			var d01 = shareholdings.Append(zeroMatrix).Append(zeroMatrix);
			var d02 = outside.Append(id).Append(zeroMatrix);
			var d03 = remainder.Append(zeroMatrix).Append(id);

			f = d01.Transpose ().Append (d02.Transpose ()).Append (d03.Transpose ()).Transpose ();

			d0 = dividends.Transpose ().Append (zeroVector).Append (zeroVector).Transpose ();

			dynamicDividend = DividendsCalc.DividendsCalc.dynamicDividendTable (d0, f);
		}

		[Test]
		public void CorrectOwnershipTable ()
		{
			var res = DividendsCalc.DividendsCalc.ownershipTable (f);
			Assert.That (res, Is.EqualTo (ownership).Using (new MatrixWithinComparer(10e-8)));
		}

		[Test]
		public void CorrectPenultimateExit ()
		{
			var res = DividendsCalc.DividendsCalc.penultimateExitTable (dynamicDividend, shareholdings, outside, remainder);
			Assert.That (res, Is.EqualTo (penultimateExit).Using (new MatrixWithinComparer(10e-8)));
		}
	}
}
