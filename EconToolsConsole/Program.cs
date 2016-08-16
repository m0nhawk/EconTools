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

			string inputFilename = string.Empty;
			string outputFilename = string.Empty;

			var option_set = new OptionSet () {
				{ "i|input=", "filename with input dividend matrix.",
					v => inputFilename = v },
				{ "o|output=", "filename for resulting tables.",
					v => outputFilename = v},
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
					"EconTools.exe [-i|--input=<filename>] [-o|--output=<filename>] [--help]";
				ShowHelp (usage_message, option_set);
			}

			if (!String.IsNullOrEmpty(inputFilename)) {
				if (File.Exists (inputFilename)) {
					switch (Path.GetExtension (inputFilename)) {
					case ".xlsx":
						var vals = Helpers.Helpers.readData (inputFilename);

						var val = Helpers.Helpers.ConvertTo (vals);

						var f = val.Item1;
						var d0 = val.Item2;

						var dtable = DividendFlowCalculator.DividendFlowCalculator.dynamicDividendTable (d0, f);

						var s = DividendFlowCalculator.DividendFlowCalculator.ownershipTable (f);

						var eTable = DividendFlowCalculator.DividendFlowCalculator.penultimateExitTable (dtable, vals.Item1, vals.Item2, vals.Item3);

						bool res;
						var companies = vals.Item5;

						if (String.IsNullOrEmpty (outputFilename)) {
							res = Helpers.Helpers.writeData (inputFilename, s, dtable.Item1, eTable, companies);
						} else {
							res = Helpers.Helpers.writeData (outputFilename, s, dtable.Item1, eTable, companies);
						}
						break;
					default:
						break;
					}
				}
			}
		}
	}
}
