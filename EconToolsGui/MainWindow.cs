using System;
using Gtk;
using MathNet.Numerics.LinearAlgebra;

using Helpers;
using DividendsCalc;

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
	static int n;

	Matrix<double> c, o, r, g;

	private void buttonOpenClicked (object sender, EventArgs e)
	{
		try
			{
			Gtk.FileChooserDialog dlg =
				new Gtk.FileChooserDialog("Choose the file to open",
					this,
					FileChooserAction.Open,
					"Cancel",ResponseType.Cancel,
					"Open",ResponseType.Accept);
			//dlg.DefaultExt = ".xlsx";
			//dlg.Filter = "Excel Files (*.xlsx)|*.xlsx";

			if (dlg.Run() == (int)ResponseType.Accept) 
			{
				string filename = dlg.Filename;
				entryInput.Text = filename;

				switch(System.IO.Path.GetExtension(filename))
				{
				case ".xlsx":
					var vals = Helpers.Helpers.readData(filename);

					c = vals.Item1;
					n = c.ColumnCount;

					o = vals.Item2;
					r = vals.Item3;
					g = vals.Item4;
					break;
				default:
					break;
				}
			}

			dlg.Destroy();
		}
		catch(Exception ex) {
			MessageDialog msg= new MessageDialog(this, 
				DialogFlags.DestroyWithParent, MessageType.Error, 
				ButtonsType.Close, ex.Message);
			msg.Title="Error";
			ResponseType response = (ResponseType) msg.Run();
			if (response == ResponseType.Close || response == ResponseType.DeleteEvent) {
				msg.Destroy();
			}
		}
	}

	private void buttonCalcClicked (object sender, EventArgs e)
	{
		try
		{
			Matrix<double> zeroVector = Matrix<double>.Build.Dense (1, n, 0.0);
			Matrix<double> zeroMatrix = Matrix<double>.Build.Dense (n, n, 0.0);
			Matrix<double> id = Matrix<double>.Build.DenseIdentity (n);

			var d01 = c.Append(zeroMatrix).Append(zeroMatrix);
			var d02 = o.Append(id).Append(zeroMatrix);
			var d03 = r.Append(zeroMatrix).Append(id);

			var f = d01.Transpose().Append(d02.Transpose()).Append(d03.Transpose()).Transpose();

			var d0 = g.Transpose().Append(zeroVector).Append(zeroVector).Transpose();

			var dtable = DividendsCalc.DividendsCalc.dynamicDividendTable(d0, f);

			var s = DividendsCalc.DividendsCalc.ownershipTable(f);

			var eTable = DividendsCalc.DividendsCalc.penultimateExitTable(dtable, c, o, r);

			Helpers.Helpers.writeData(entryInput.Text, s, dtable.Item1, eTable);
		}
		catch(Exception ex) {
			MessageDialog msg= new MessageDialog(this, 
				DialogFlags.DestroyWithParent, MessageType.Error, 
				ButtonsType.Close, ex.Message);
			msg.Title="Error";
			ResponseType response = (ResponseType) msg.Run();
			if (response == ResponseType.Close || response == ResponseType.DeleteEvent) {
				msg.Destroy();
			}
		}
	}
}
