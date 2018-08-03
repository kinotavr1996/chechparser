using ParsingSystem.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ParsingSystem.Proccessor
{
	public class ExcelParserProccessor
	{
		private static Excel.Workbook MyBook = null;
		private static Excel.Application MyApp = null;
		private static Excel.Worksheet MySheet = null;
		private int LastRow = 0;
		private int LastColumn;
		private string Path;
		public void Open(string path)
		{
			if (string.IsNullOrEmpty(path))
				return;
			Path = path;
			MyApp = new Excel.Application
			{
				Visible = false
			};
			MyBook = MyApp.Workbooks.Open(path);
			MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here
			var info = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
			LastRow = info.Row;
			LastColumn = info.Column;
		}
		public Excel.Application Create()
		{
			var app = new Excel.Application();
			object misValue = System.Reflection.Missing.Value;
			app.Workbooks.Add(misValue);
			return app;
		}
		public Excel.Application Copy(Excel.Application file)
		{
			Excel.Application copy = new Excel.Application();
			Excel.Workbook xlWb = file.ActiveWorkbook as Excel.Workbook;
			Excel.Worksheet xlSht = xlWb.Sheets[1];
			xlSht.Copy(Type.Missing, xlWb.Sheets[xlWb.Sheets.Count]); // copy
			xlWb.Sheets[xlWb.Sheets.Count].Name = file.ActiveWorkbook.Name;
			copy.Workbooks.Add(xlWb);
			return copy;
		}
		public void Save(Excel.Workbook file = null, string path = null, string name = null)
		{
			path = path ?? Path;
			name = name ?? MyBook.Name;
			file = file ?? MyBook;
			file.SaveAs($"{path}/{name}.xls", Excel.XlFileFormat.xlWorkbookNormal);
		}
		public List<ProductInfo> Read()
		{
			var productList = new List<ProductInfo>();
			for (int index = 2; index <= LastRow; index++)
			{
				var MyValues = (Array)MySheet.get_Range("A" +
				   index.ToString(), "AI" + index.ToString()).Cells.Value;
				int.TryParse(MyValues.GetValue(1, 7).ToString(), out int itemId);
				decimal.TryParse(MyValues.GetValue(1, 12).ToString(), out decimal yourPrice);
				productList.Add(new ProductInfo
				{
					ItemId = itemId,
					Url = MyValues.GetValue(1, 9).ToString(),
					YourPrice = yourPrice
				});
			}
			return productList;
		}
		public void Write(List<ProductInfo> productList)
		{
			//var copy = Copy(MyApp);
			//var sheet = (Excel.Worksheet)copy.Sheets[1];
			for (int i = 2; i <= LastRow; i++)
			{
				var flag = int.TryParse(MySheet.Cells[i, 7].Value, out int itemId);
				if (!flag) continue;
				var product = productList.FirstOrDefault(x => x.ItemId == itemId);
				MySheet.Cells[i, 16] = product.LowestPrice;
				var j = 17;
				foreach(var price in product.Prices.OrderBy(x => x))
				{
					MySheet.Cells[i, j] = price.ToString();
					j++;
				}
				
			}

			Save(name: $"{MyBook.Name}_output");
		}
	}
}
