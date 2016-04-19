using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IESConverter
{
	class Program
	{
		[STAThread]
		static void Main(string[] args)
		{
			var iesFilePath = "";

			Console.WriteLine("Choose IPF file");

			OpenFileDialog ofd = new OpenFileDialog();
			ofd.Filter = "IES Files| *.ies";
			if (ofd.ShowDialog() == DialogResult.OK)
				iesFilePath = ofd.FileName;

			var file = File.ReadAllBytes(iesFilePath);
			var fileName = Path.GetFileName(iesFilePath);

			new MakeExcel(new IesFile(file), fileName.Substring(0, fileName.LastIndexOf('.')), System.IO.Path.GetDirectoryName(iesFilePath));
			
			Console.Read();
		}
	}
}
