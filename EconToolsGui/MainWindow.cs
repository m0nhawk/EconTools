using System;
using Gtk;
using MathNet.Numerics.LinearAlgebra;

using DividendFlowCalculator;

public sealed partial class MainWindow: Gtk.Window
{
	public MainWindow () : base (Gtk.WindowType.Toplevel)
	{
		Build ();
	}

	private void OnDeleteEvent (object sender, DeleteEventArgs a)
	{
		Application.Quit ();
		a.RetVal = true;
	}

	private Tuple<Matrix<double>, Matrix<double>, Matrix<double>, Matrix<double>> vals;

	private void buttonOpenClicked (object sender, EventArgs e)
	{
		try {
			Gtk.FileChooserDialog dlg =
				new Gtk.FileChooserDialog ("Choose the file to open",
					this,
					FileChooserAction.Open,
					"Cancel", ResponseType.Cancel,
					"Open", ResponseType.Accept);
			
			var excel = new FileFilter ();
			excel.Name = "Excel Files";
			excel.AddPattern("*.xlsx");

			var csv = new FileFilter ();
			csv.Name = "Comma-Separated Values";
			csv.AddPattern("*.csv");

			dlg.AddFilter(excel);

			if (dlg.Run () == (int)ResponseType.Accept) {
				string filename = dlg.Filename;
				entryInput.Text = filename;

				switch (System.IO.Path.GetExtension (filename)) {
				case ".xlsx":
					vals = Helpers.Helpers.readData (filename);
					break;
				default:
					break;
				}
			}

			dlg.Destroy ();
		} catch (Exception ex) {
			MessageDialog msg = new MessageDialog (this, 
				                   DialogFlags.DestroyWithParent, MessageType.Error, 
				                   ButtonsType.Close, ex.Message);
			msg.Title = "Error";
			ResponseType response = (ResponseType)msg.Run ();
			if (response == ResponseType.Close || response == ResponseType.DeleteEvent) {
				msg.Destroy ();
			}
		}
	}

	private void buttonCalcClicked (object sender, EventArgs e)
	{
		try {
			var val = Helpers.Helpers.ConvertTo (vals);

			var f = val.Item1;
			var d0 = val.Item2;

			var dtable = DividendFlowCalculator.DividendFlowCalculator.dynamicDividendTable (d0, f);

			var s = DividendFlowCalculator.DividendFlowCalculator.ownershipTable (f);

			var eTable = DividendFlowCalculator.DividendFlowCalculator.penultimateExitTable (dtable, vals.Item1, vals.Item2, vals.Item3);

			Helpers.Helpers.writeData (entryInput.Text, s, dtable.Item1, eTable);
		} catch (Exception ex) {
			MessageDialog msg = new MessageDialog (this, 
				                   DialogFlags.DestroyWithParent, MessageType.Error, 
				                   ButtonsType.Close, ex.Message);
			msg.Title = "Error";
			ResponseType response = (ResponseType)msg.Run ();
			if (response == ResponseType.Close || response == ResponseType.DeleteEvent) {
				msg.Destroy ();
			}
		}
	}
}
