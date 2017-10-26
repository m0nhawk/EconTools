using NUnit.Framework;
using System;
using System.Collections.Generic;
using MathNet.Numerics.LinearAlgebra;

using DividendFlowCalculator;
using System.IO;

namespace GivenAData
{
	public class MatrixWithinComparer : IComparer<Matrix<double>>
	{
		private double epsilon;

		public MatrixWithinComparer (double epsilon = 10e-6)
		{
			this.epsilon = epsilon;
		}

		int IComparer<Matrix<double>>.Compare (Matrix<double> a, Matrix<double> b)
		{
			foreach (var v in a.EnumerateIndexed ()) {
				if (v.Item3 - b.At (v.Item1, v.Item2) > epsilon)
					return -1;
			}
			return 0;
		}
	}

	[TestFixture ()]
	public class OnCalculation
	{
		private Matrix<double> shareholdings, outside, remainder, dividends;
		private DividendData data;
		private Matrix<double> dynamicDividend;
		private Matrix<double> ownership;
		private Matrix<double> penultimateExit;

		private int length = 3;

		[SetUp ()]
		public void Init ()
		{
			shareholdings = Matrix<double>.Build.DenseOfArray (new double[,] {
				{ 0, 0.1, 0.2 },
				{ 0, 0, 0.15 },
				{ 0.12, 0.12, 0 }
			});
			
			outside = Matrix<double>.Build.DenseOfDiagonalArray (new double[] { 0.11, 0.21, 0.13 });

			remainder = Matrix<double>.Build.DenseOfDiagonalArray (new double[] { 0.77, 0.57, 0.52 });

			dividends = Matrix<double>.Build.DenseOfColumnMajor (length, 1, new double[] { 1000, 0, 0 });

			ownership = Matrix<double>.Build.DenseOfArray (new double[,] {
				{ 0.1129679439, 0.0142647149, 0.0247332029 },
				{ 0.0039530958, 0.2143483761, 0.0329427647 },
				{ 0.016314512, 0.0179459484, 0.1359547246 },
				{ 0.7907756072, 0.099853004, 0.17313242 },
				{ 0.0107298313, 0.5818027351, 0.0894160756 },
				{ 0.0652580481, 0.0717837935, 0.5438188982 }
			});

			penultimateExit = Matrix<double>.Build.DenseOfArray (new double[,] {
				{ 110, 0.2070696487, 2.7609286707 },
				{ 0, 0, 3.9531478695 },
				{ 16.0209161215, 0.2936624109, 0 },
				{ 770, 1.4494875412, 19.3265006951 },
				{ 0, 0, 10.7299727885 },
				{ 64.0836644861, 1.1746496438, 0 }
			});

			dynamicDividend = Matrix<double>.Build.DenseOfArray(new double[,] {
				{ 1000, 0, 24, 1.8, 1.008, 0.1188, 0.045576, 0.006804, 0.002128032, 0.0003678048, 0.0001016245, 1.92782592E-05, 4.930279488E-06, 9.926110656E-07 },
				{ 0, 0, 18, 0, 0.756, 0.0324, 0.031752, 0.0027216, 0.001391904, 0.0001714608, 6.3358848E-05, 9.7067808E-06, 2.969701056E-06, 5.2173072E-07 },
				{ 0, 120, 0, 5.04, 0.216, 0.21168, 0.018144, 0.00927936, 0.001143072, 0.0004223923, 6.4711872E-05, 0.000019798, 3.4782048E-06, 0.000000948 },
				{ 0, 110, 110, 112.64, 112.838, 112.94888, 112.961948, 112.96696136, 112.9677098, 112.9679438835, 112.967984342, 112.9679955207, 112.9679976414, 112.9679981837 },
				{ 0, 0, 0, 3.78, 3.78, 3.93876, 3.945564, 3.95223192, 3.952803456, 3.9530957558, 3.9531317626, 3.953145068, 3.9531471064, 3.95314773 },
				{ 0, 0, 15.6, 15.6, 16.2552, 16.28328, 16.3107984, 16.31315712, 16.3143634368, 16.3145120362, 16.3145669472, 16.3145753597, 16.3145779334, 16.3145783856 },
				{ 0, 770, 770, 788.48, 789.866, 790.64216, 790.733636, 790.76872952, 790.7739686, 790.7756071846, 790.7758903943, 790.7759686452, 790.7759834895, 790.7759872858 },
				{ 0, 0, 0, 10.26, 10.26, 10.69092, 10.709388, 10.72748664, 10.729037952, 10.7298313373, 10.7299290699, 10.7299651845, 10.7299707173, 10.7299724101 },
				{ 0, 0, 62.4, 62.4, 65.0208, 65.13312, 65.2431936, 65.25262848, 65.2574537472, 65.2580481446, 65.2582677886, 65.2583014388, 65.2583117338, 65.2583135425 }
			});

			data = new DividendFlowCalculator.DividendData ();
			data.shareholdings = shareholdings;
			data.outside = outside;
			data.remainder = remainder;
			data.dividends = dividends;
			data.companies = new List<string> ();
		}

		[Test ()]
		public void CorrectDynamicDidivend ()
		{
			var dynamicDividendResult = data.dynamicDividend;
			Assert.That (dynamicDividendResult, Is.EqualTo (dynamicDividend).Using (new MatrixWithinComparer ()));
		}

		[Test ()]
		public void CorrectOwnershipTable ()
		{
			var ownershipResult = data.ownership;
			Assert.That (ownershipResult, Is.EqualTo (ownership).Using (new MatrixWithinComparer ()));
		}

		[Test ()]
		public void CorrectPenultimateExit ()
		{
			var penultimateExitResult = data.penultimateExit;
			Assert.That (penultimateExitResult, Is.EqualTo (penultimateExit).Using (new MatrixWithinComparer ()));
		}
	}
}

namespace GivenAFile
{
	[TestFixture ()]
	public class OnLoad
	{
		private Matrix<double> shareholdingsOk, outsideOk, remainderOk, dividendsOk;

		[SetUp ()]
		public void Init ()
		{
			int length = 3;

			shareholdingsOk = Matrix<double>.Build.DenseOfArray (new double[,] { { 0, 0.1, 0.2 },
				{ 0, 0, 0.15 },
				{ 0.12, 0.12, 0 }
			});
			outsideOk = Matrix<double>.Build.DenseOfDiagonalArray (new double[] { 0.11, 0.21, 0.13 });
			remainderOk = Matrix<double>.Build.DenseOfDiagonalArray (new double[] { 0.77, 0.57, 0.52 });
			dividendsOk = Matrix<double>.Build.DenseOfColumnMajor (length, 1, new double[] { 1000, 0, 0 });
		}

        [Test()]
        public void Success()
        {
            var data = new DividendFlowCalculator.DividendData();

            data.LoadFromFile(Path.Combine(TestContext.CurrentContext.TestDirectory, "success.xlsx"));

            Matrix<double> readShareholdings, readOutside, readRemainder, readDividends;

            readShareholdings = data.shareholdings;
            readOutside = data.outside;
            readRemainder = data.remainder;
            readDividends = data.dividends;

            Assert.AreEqual(readShareholdings, shareholdingsOk);
            Assert.AreEqual(readOutside, outsideOk);
            Assert.AreEqual(readRemainder, remainderOk);
            Assert.AreEqual(readDividends, dividendsOk);
        }

        [Test()]
        public void MissingFile()
        {
            var data = new DividendFlowCalculator.DividendData();
            Assert.Throws<System.IO.FileNotFoundException>(() => data.LoadFromFile(Path.Combine(TestContext.CurrentContext.TestDirectory, "missing.xlsx")));
        }

        [Test()]
        public void InvalidData()
        {
            var data = new DividendFlowCalculator.DividendData();
            var notOk = data.LoadFromFile(Path.Combine(TestContext.CurrentContext.TestDirectory, "invalid.xlsx"));
            Assert.IsFalse(notOk);
        }
    }
}