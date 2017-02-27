using System;
using System.Collections.Generic;
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

	private DividendData data;

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
					data = new DividendFlowCalculator.DividendData ();
					data.LoadFromFile (filename);
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
			data.writeData (entryInput.Text);
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
