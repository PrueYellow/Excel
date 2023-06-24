using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Excel
{
	internal class ReadExcel
	{

		public ReadExcel()

		{
			//放EXCEL檔路徑
			string sourcePath = @"C:\Users\G-pro\Downloads\BookProducts.xlsx";

			Microsoft.Office.Interop.Excel.Application _application
				= new Microsoft.Office.Interop.Excel.Application();

			_Workbook book = _application.Workbooks.Open(sourcePath);

			_Worksheet sheet;
			try
			{
				sheet = book.Worksheets.get_Item(1);

				Console.WriteLine(sheet);

				int row = sheet.UsedRange.Rows.Count;
				int columns = sheet.UsedRange.Columns.Count;
				
				Range range = sheet.Range[sheet.Cells[1,1],sheet.Cells[row,columns]];

				Array result = range.Value2;

				for(int i = 2; i <= columns; i++)
				{
					for(int j = 1; j <= row; j++)
					{
						Console.WriteLine((string)result.GetValue(1, j)+" " + result.GetValue(i, j));
					}
				}



				Console.ReadKey();
			}
			catch (Exception ex)
			{
				Console.WriteLine("An error occurred: " + ex.Message);

			}
		}

		static void Main(string[] args)
		{
		
			ReadExcel excel = new ReadExcel();

			Console.WriteLine(excel);
			Console.ReadKey();
		}
	}
}
