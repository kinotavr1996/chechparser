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
		public MainWindow()
		{
			InitializeComponent();
		}

		private void btnOpenFile_Click(object sender, RoutedEventArgs e)
		{
			//var openFileDialog = new OpenFileDialog();
			//if (openFileDialog.ShowDialog() == true)
			//	txtEditorBrowse.Text = File.ReadAllText(openFileDialog.FileName);
			var tempUrl = "C:\\Users\\admin\\Downloads\\AAA.xlsx";
			//var productList = Load(openFileDialog.FileName);
			var productList = Load(tempUrl);
			var priceList = new List<string>();
			foreach(var item in productList)
			{
				var doc = new HtmlWeb().Load(item.Url);
				var elements = doc.DocumentNode.SelectNodes("//*[@class=\"pricen\"]");
				priceList.AddRange(elements.Select(x => string.IsNullOrEmpty(x.InnerText.ToString()) ?
										x.InnerText.ToString() :
										x.InnerHtml.ToString()
								));
			}

		}

		private void btnSaveFile_Click(object sender, RoutedEventArgs e)
		{
			var dlg = new SaveFileDialog
			{
				FileName = "Result", // Default file name
				DefaultExt = ".text", // Default file extension
				Filter = "Text documents (.txt)|*.txt" // Filter files by extension
			};

			// Show save file dialog box
			Nullable<bool> result = dlg.ShowDialog();

			// Process save file dialog box results
			if (result == true)
			{
				// Save document
				string filename = dlg.FileName;
			}
		}

		private void btnConfigure_Click(object sender, RoutedEventArgs e)
		{

		}

		private void btnScan_Click(object sender, RoutedEventArgs e)
		{

		}

		private void btnSettings_Click(object sender, RoutedEventArgs e)
		{

		}

		private List<ProductInfo> Load(string path)
		{
			var excelProccessor = new ExcelParserProccessor();
			excelProccessor.Open(path);
			var productList = excelProccessor.Read();
			return productList;
		} 
	}
}
