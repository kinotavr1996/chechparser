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
			if (MyBook != null) MyBook.Close(0);
			if (MyApp != null) MyApp.Quit();
		}
		public List<ProductInfo> ReadSheetBySheetIndex(int sheetIndex)
		{
			var productList = new List<ProductInfo>();
			if (sheetIndex > MyApp.Sheets.Count) return productList;
			UpdateCurrentSheetInfo(sheetIndex);
			for (int index = 20; index <= LastRow; index++)
			{
				var MyValues = (Array)MySheet.get_Range("A" +
				   index.ToString(), "G" + index.ToString()).Cells.Value;
				int.TryParse(MyValues.GetValue(1, 3)?.ToString(), out int itemId);
				decimal.TryParse(MyValues.GetValue(1, 6)?.ToString(), out decimal yourPrice);
				decimal.TryParse(MyValues.GetValue(1, 7)?.ToString(), out decimal postageCost);
				productList.Add(new ProductInfo
				{
					Category = MyValues.GetValue(1, 1)?.ToString() ?? string.Empty,
					ItemId = itemId,
					Description = MyValues.GetValue(1, 4)?.ToString() ?? string.Empty,
					Url = MyValues.GetValue(1, 5)?.ToString() ?? string.Empty,
					Price = yourPrice,
					PostageCost = postageCost,
					LowestPrice = yourPrice
				});
			}
			return productList;
		}
		public List<ProductInfo> Read()
		{
			var productList = new List<ProductInfo>();
			AddIfNotExistMasterSheet();
			for (int i = 2; i <= MyApp.Sheets.Count; i++)
			{
				var data = ReadSheetBySheetIndex(i);
				if (data != null) productList.AddRange(data);
			}
			return productList;
		}

		public void Write(List<ProductInfo> productList)
		{
			if (productList?.Count == 0) return;
			for (int i = 2; i <= MyApp.Sheets.Count; i++)
			{
				UpdateCurrentSheetInfo(i);
				for (int j = 20; j <= LastRow; j++)
				{
					var MyValues = (Array)MySheet.get_Range("A" + j.ToString(), "G" + j.ToString()).Cells.Value;
					var flag = int.TryParse(MyValues.GetValue(1, 3)?.ToString(), out int itemId);
					if (!flag) continue;
					var product = productList.FirstOrDefault(x => x.ItemId == itemId);
					if (product == null) continue;
					if (product.Price != product.LowestPrice)
					{
						MySheet.Cells[j, 6].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
						MySheet.Cells[j, 6] = product.Price;
						MySheet.Cells[j, 7] = product.LowestPrice;
						MySheet.Cells[j, 7].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
					}
				}
			}
			WriteInfoToMasterSheet(productList);
			//Name = $"{DateTime.Now.ToString("yyyyMMddHHmmss")}--{MyBook.Name}";
		}

		private void AddIfNotExistMasterSheet()
		{
			foreach (Excel.Worksheet sheet in MyApp.Sheets)
				if (sheet.Name == "Master Sheet") return;

			var xlNewSheet = (Excel.Worksheet)MyApp.Sheets.Add(MyApp.Sheets[1], Type.Missing, Type.Missing, Type.Missing);
			xlNewSheet.Name = "Master Sheet";
		}

		private void UpdateCurrentSheetInfo(int sheetIndex)
		{
			MySheet = (Excel.Worksheet)MyBook.Sheets[sheetIndex]; // Explicit cast is not required here
			var info = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
			LastRow = info.Row;
			LastColumn = info.Column;
		}

		private void WriteInfoToMasterSheet(List<ProductInfo> productList)
		{
			UpdateCurrentSheetInfo(1);
			for (int j = 20; j <= productList.Count; j++)
			{
				var product = productList[j - 20];
				if (!product.IsParsed || product.Price != product.LowestPrice)
				{
					MySheet.Cells[j, 1] = product.Category;
					MySheet.Cells[j, 3] = product.ItemId;
					MySheet.Cells[j, 4] = product.Description;
					MySheet.Cells[j, 5] = product.Url;
					if(!product.IsParsed)
						MySheet.Cells[j, 5].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
					MySheet.Cells[j, 6] = product.Price;
					MySheet.Cells[j, 6].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
					MySheet.Cells[j, 7] = product.LowestPrice;
					MySheet.Cells[j, 7].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
				}
			}
		}
	}
}
