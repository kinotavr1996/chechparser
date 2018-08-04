using ParsingSystem.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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
		public string Name { get; set; }
		public void Open(string path)
		{
			if (string.IsNullOrEmpty(path))
				return;
			MyApp = new Excel.Application
			{
				Visible = false
			};
			MyBook = MyApp.Workbooks.Open(path);
			
		}
		private void Initialize()
		{
			MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here
			var info = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
			LastRow = info.Row;
			LastColumn = info.Column;
			Name = MyBook.Name;
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
			path = path ?? MyBook.Path;
			name = name ?? Name;
			file = file ?? MyBook;
			file.SaveAs($"{path}\\{name}", Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
					Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
					Excel.XlSaveConflictResolution.xlUserResolution, true,
					Missing.Value, Missing.Value, Missing.Value);
			Quit();
		}
		public void Quit()
		{
			MyBook.Close(0);
			MyApp.Quit();
		}
		public List<ProductInfo> ReadSheetBySheetIndex(int sheetIndex)
		{
			var productList = new List<ProductInfo>();

			if (sheetIndex > MyApp.Sheets.Count) return productList;
			MySheet = (Excel.Worksheet)MyBook.Sheets[sheetIndex]; // Explicit cast is not required here
			for (int index = 2; index <= LastRow; index++)
			{
				var MyValues = (Array)MySheet.get_Range("A" +
				   index.ToString(), "AI" + index.ToString()).Cells.Value;
				int.TryParse(MyValues.GetValue(1, 7)?.ToString(), out int itemId);
				decimal.TryParse(MyValues.GetValue(1, 12)?.ToString(), out decimal yourPrice);
				productList.Add(new ProductInfo
				{
					ItemId = itemId,
					Url = MyValues.GetValue(1, 9)?.ToString() ?? string.Empty,
					YourPrice = yourPrice,
					Prices = new List<decimal>()
				});
			}
			return productList;
		}
		public List<ProductInfo> Read()
		{
			var productList = new List<ProductInfo>();
			for (int i = 1; i <= MyApp.Sheets.Count; i++)
			{
				var data = ReadSheetBySheetIndex(i);
				if (data != null) productList.AddRange(data);
			}
			return productList;
		}

		public void Write(List<ProductInfo> productList)
		{
			for (int i = 2; i <= LastRow; i++)
			{
				var flag = int.TryParse(MySheet.Cells[i, 7].Value, out int itemId);
				if (!flag) continue;
				var product = productList.FirstOrDefault(x => x.ItemId == itemId);
				MySheet.Cells[i, 16] = product.LowestPrice;
				var j = 17;
				foreach (var price in product.Prices.OrderBy(x => x))
				{
					MySheet.Cells[i, j] = price.ToString();
					j++;
				}

			}
			Name = $"{DateTime.Now.ToString("yyyyMMddHHmmss")}--{MyBook.Name}";
		}

	}
}
