using System;
using Eto.Forms;

namespace EconTools
{
	class MainClass
	{
        [STAThread]
        public static void Main (string[] args)
		{
            new Application().Run(new MyForm());
        }
	}
}
