using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace IESConverter
{
	class MakeExcel
	{
		public MakeExcel(IesFile table, string name, string outputPath)
		{
			Console.WriteLine("Starting excel");
			var xlApp = new Application();
			var misValue = System.Reflection.Missing.Value;
			var xlWorkBook = xlApp.Workbooks.Add(misValue);
			var xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.Item[1];
			xlWorkSheet.Name = name;

			for (var i = 0; i < table.Columns.Count; i++)
				xlWorkSheet.Cells[1, i+1] = table.Columns[i].Name;
			for (var i = 0; i < table.Rows.Count; i++)
			{
				Console.WriteLine($"{i} / {table.Rows.Count} done");
				var rowValues = table.Rows[i].Values.ToList();
				for (var j = 0; j < rowValues.Count; j++)
					xlWorkSheet.Cells[i + 2, j + 1] = rowValues[j];
			}

			xlWorkBook.SaveAs(outputPath + $"\\{name}.xls", XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
			xlWorkBook.Close(true, misValue, misValue);
			xlApp.Quit();

			Console.WriteLine($"Done. {outputPath}\\{name}.xls");
		}
	}
}
