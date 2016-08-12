using System;
using File = System.IO.File;
using Path = System.IO.Path;
using MathNet.Numerics.LinearAlgebra;

using DividendFlowCalculator;

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
				switch (Path.GetExtension (filename)) {
				case ".xlsx":
					var vals = Helpers.Helpers.readData (filename);

					var val = Helpers.Helpers.ConvertTo (vals);

					var f = val.Item1;
					var d0 = val.Item2;

					var dtable = DividendFlowCalculator.DividendFlowCalculator.dynamicDividendTable (d0, f);

					var s = DividendFlowCalculator.DividendFlowCalculator.ownershipTable (f);

					var eTable = DividendFlowCalculator.DividendFlowCalculator.penultimateExitTable (dtable, vals.Item1, vals.Item2, vals.Item3);

					Helpers.Helpers.writeData (filename, s, dtable.Item1, eTable);

					break;
				default:
					break;
				}
			}
		}
	}
}
