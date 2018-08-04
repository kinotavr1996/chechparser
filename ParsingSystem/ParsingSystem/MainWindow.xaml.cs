using HtmlAgilityPack;
using Microsoft.Win32;
using ParsingSystem.Models;
using ParsingSystem.Proccessor;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ParsingSystem
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private readonly ExcelParserProccessor excelProccessor = new ExcelParserProccessor();

		private readonly OpenFileDialog openFileDialog = new OpenFileDialog();
		public MainWindow()
		{
			InitializeComponent();
		}

		private void btnOpenFile_Click(object sender, RoutedEventArgs e)
		{
			if (openFileDialog.ShowDialog() == true)
				txtEditorBrowse.Text = File.ReadAllText(openFileDialog.FileName);
		}

		private void btnSaveFile_Click(object sender, RoutedEventArgs e)
		{
			var dlg = new SaveFileDialog
			{
				FileName = excelProccessor.Name, // Default file name
				DefaultExt = ".xlsx", // Default file extension
				Filter = "Excel documents (.xlsx)|*.xlsx" // Filter files by extension
			};
			Nullable<bool> result = dlg.ShowDialog();
			if (result != true) return;
			// Save document
			excelProccessor.Name = dlg.SafeFileName;
			excelProccessor.Save();
		}

		private void btnConfigure_Click(object sender, RoutedEventArgs e)
		{

		}

		private void btnScan_Click(object sender, RoutedEventArgs e)
		{
			Parse();
		}

		private void btnSettings_Click(object sender, RoutedEventArgs e)
		{

		}

		private List<ProductInfo> Load(string path)
		{
			excelProccessor.Open(path);
			var productList = excelProccessor.ReadFromSheetBySheetIndex();
			return productList;
		}

		private void Parse()
		{
			try
			{
				var productList = Load(openFileDialog.FileName);
				var priceList = new List<string>();
				foreach (var item in productList)
				{
					if (string.IsNullOrEmpty(item.Url)) continue;
					var doc = new HtmlWeb().Load(item.Url);
					var elements = doc.DocumentNode.SelectNodes("//*[@class=\"single-price\"]");
					if (elements == null) continue;
					foreach (var el in elements)
					{
						var price = new string((string.IsNullOrEmpty(el.InnerText.ToString()) ?
									el.InnerText.Where(x => char.IsDigit(x)) :
									el.InnerHtml.Where(x => char.IsDigit(x))).ToArray());
						decimal.TryParse(price, out decimal parsedPrice);
						item.Prices.Add(parsedPrice);
					}
					item.Prices = item.Prices.OrderBy(x => x).ToList();
					item.LowestPrice = item.Prices.FirstOrDefault();
					item.Prices.Remove(item.LowestPrice);
				}

				excelProccessor.Write(productList);
			}
			catch (Exception ex)
			{
				Console.Write(ex);
				excelProccessor.Quit();
			}
		}
	}
}
