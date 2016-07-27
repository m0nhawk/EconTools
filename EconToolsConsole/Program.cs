using System;
using File = System.IO.File;
using Path = System.IO.Path;
using MathNet.Numerics.LinearAlgebra;

using Helpers;
using DividendsCalc;

namespace EconToolsConsole
{
	class MainClass
	{
		public static void Main (string[] args)
		{
			string filename = "";
			if (args.Length == 1) {
				filename = args [0];
			}

			if (File.Exists (filename)) {
				Matrix<double> c, o, r, g;
				int n = 0;

				switch (Path.GetExtension (filename)) {
				case ".xlsx":
					var vals = Helpers.Helpers.readData (filename);

					c = vals.Item1;
					n = c.ColumnCount;

					o = vals.Item2;
					r = vals.Item3;
					g = vals.Item4;

					Matrix<double> zeroVector = Matrix<double>.Build.Dense (1, n, 0.0);
					Matrix<double> zeroMatrix = Matrix<double>.Build.Dense (n, n, 0.0);
					Matrix<double> id = Matrix<double>.Build.DenseIdentity (n);

					var d01 = c.Append(zeroMatrix).Append(zeroMatrix);
					var d02 = o.Append(id).Append(zeroMatrix);
					var d03 = r.Append(zeroMatrix).Append(id);

					var f = d01.Transpose ().Append (d02.Transpose ()).Append (d03.Transpose ()).Transpose ();

					var d0 = g.Transpose ().Append (zeroVector).Append (zeroVector).Transpose ();

					var dtable = DividendsCalc.DividendsCalc.dynamicDividendTable (d0, f);

					var s = DividendsCalc.DividendsCalc.ownershipTable (f);

					var eTable = DividendsCalc.DividendsCalc.penultimateExitTable (dtable, c, o, r);

					Helpers.Helpers.writeData (filename, s, dtable.Item1, eTable);

					break;
				default:
					break;
				}
			}
		}
	}
}
