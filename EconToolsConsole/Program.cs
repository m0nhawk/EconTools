using System;
using System.Collections.Generic;
using Mono.Options;
using File = System.IO.File;
using Path = System.IO.Path;
using MathNet.Numerics.LinearAlgebra;

using DividendFlowCalculator;

namespace EconToolsConsole
{
	class MainClass
	{
		public static void ShowHelp(string message, OptionSet option_set)
		{
			Console.Error.WriteLine(message);
			option_set.WriteOptionDescriptions(Console.Error);
			Environment.Exit(-1);
		}

		public static void Main (string[] args)
		{
			bool help = false;

			string filename = string.Empty;

			var option_set = new OptionSet () {
				{ "o|open=", "filename with input dividend matrix.",
					v => filename = v },
				{ "h|help",  "help message.", 
					v => help = v != null },
			};

			List<string> extra;
			try {
				extra = option_set.Parse (args);
			}
			catch (OptionException e) {
				Console.Write ("EconTools.exe: ");
				Console.WriteLine (e.Message);
				Console.WriteLine ("Try `EconTools.exe --help' for more information.");
			}

			if (help) {
				const string usage_message = 
					"EconTools.exe [--o|open=<filename>] [--help]";
				ShowHelp (usage_message, option_set);
			}

			if (!String.IsNullOrEmpty(filename)) {
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
}
