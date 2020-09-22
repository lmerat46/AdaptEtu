using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace AdaptEtu
{

	public class Excel
	{
		string path = "";
		_Application excel = new _Excel.Application();
		Workbook wb;
		Worksheet ws;

		public Excel(string path, int sheet)
		{
			this.path = path;
			wb = excel.Workbooks.Open(path);
			ws = wb.Worksheets[sheet];
		}

		public void read_cell(int i, int j)
		{
			i++;
			j++;
			if (ws.Cells[i, j].Value2 != null)
				Console.WriteLine(ws.Cells[i, j].Value2);
			else
				Console.WriteLine("oupsi, not working yet");
		}
	}

}
