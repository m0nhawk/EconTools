using System;
using System.Collections.Generic;
using Eto.Forms;
using Eto.Drawing;
using MathNet.Numerics.LinearAlgebra;

using DividendFlowCalculator;

public class MyForm : Form
{
    public MyForm()
    {
        string filename = null;
        DividendData dd = new DividendData();
        
        ClientSize = new Size(300, 100);

        Resizable = false;

        var filePath = new TextBox() { ReadOnly = true };

        var button = new Button();
        button.Text = "Open";
        button.Click += (sender, e) => {
            var dialog = new OpenFileDialog() { Title = "!" };
            dialog.ShowDialog(this);
            if (dialog.CheckFileExists)
            {
                filename = dialog.FileName;
                filePath.Text = filename;
                dd.LoadFromFile(filename);
            }
        };

        var button_save = new Button();
        button_save.Text = "Save";
        button_save.Click += (sender, e) =>
        {
            if (filename != null)
            {
                dd.writeData(filename);
            }
        };
        
        Title = "EconTools";
        
        Content = new TableLayout
        {
            Spacing = new Size(5, 5),
            Padding = new Padding(10, 10, 10, 10),
            Rows =
            {
                new TableRow(
                    filePath
                ),
                new TableLayout {
                    Rows = { new TableRow(new TableCell(button, true), new TableCell(button_save, true)) }
                }
            }
        };
    }
}
