using HtmlAgilityPack;
using Microsoft.Win32;
using ParsingSystem.Models;
using ParsingSystem.Proccessor;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Threading;

namespace ParsingSystem
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private readonly ExcelParserProccessor excelProccessor = new ExcelParserProccessor();
        private readonly MailingProcessor mailProccessor = new MailingProcessor();

        private readonly DispatcherTimer dtClockTime = new DispatcherTimer();
		private readonly long ticksBlocker = new DateTime(2018, 10, 1).Ticks;
		private readonly OpenFileDialog openFileDialog = new OpenFileDialog();
		public MainWindow()
		{
			InitializeComponent();
            int.TryParse(txtEditorPeriodicity.Text, out int periodicity);

            dtClockTime.Interval = new TimeSpan(periodicity, 0, 30); //in Hour, Minutes, Second.
			dtClockTime.Tick += dtClockTime_Tick;
			if ((bool)RunInBackground.IsChecked) dtClockTime.Start();
		}
		private void dtClockTime_Tick(object sender, EventArgs e)
		{
			Parse();
			excelProccessor.Save();
            mailProccessor.Send(txtEditorMail.Text, attachmentFileName: excelProccessor.Name);
        }


		private void btnOpenFile_Click(object sender, RoutedEventArgs e)
		{
			if (DateTime.UtcNow.Ticks >= ticksBlocker) return;
			if (openFileDialog.ShowDialog() == true)
				txtEditorBrowse.Text = openFileDialog.FileName;
		}

		private void btnSaveFile_Click(object sender, RoutedEventArgs e)
		{
			if (DateTime.UtcNow.Ticks >= ticksBlocker) return;

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
			txtEditorSave.Text = dlg.SafeFileName;
			excelProccessor.Save();
		}

		private void btnScan_Click(object sender, RoutedEventArgs e)
		{
			if (DateTime.UtcNow.Ticks >= ticksBlocker) return;
			Parse();
		}

		private List<ProductInfo> Load(string path)
		{
			excelProccessor.Open(path);
			var productList = excelProccessor.Read();
			return productList;
		}

		private void Parse()
		{
			if (DateTime.UtcNow.Ticks >= ticksBlocker) return;
			if (string.IsNullOrEmpty(openFileDialog.FileName)) return;
			try
			{
				var productList = Load(openFileDialog.FileName);
				var priceList = new List<string>();
				foreach (var item in productList)
				{
					if (string.IsNullOrEmpty(item.Url) || !Uri.IsWellFormedUriString(item.Url, UriKind.Absolute)) continue;
					var doc = new HtmlWeb().Load(item.Url);
					var elements = doc.DocumentNode.SelectNodes("//*[@class=\"pricen\"]");
					if (elements == null) continue;
					foreach (var el in elements)
					{
						var price = new string((string.IsNullOrEmpty(el.InnerText.ToString()) ?
									el.InnerText.Where(x => char.IsDigit(x)) :
									el.InnerHtml.Where(x => char.IsDigit(x))).ToArray());
						decimal.TryParse(price, out decimal parsedPrice);
						if (parsedPrice < 150) continue;
						if (item.LowestPrice > 0)
							item.LowestPrice = parsedPrice < item.LowestPrice ? parsedPrice : item.LowestPrice;
						else
							item.LowestPrice = parsedPrice;

						item.IsParsed = true;
					}
				}
				excelProccessor.Write(productList);
				MessageBox.Show("Parsing completed successfully", "Notification");
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "Exception");
				excelProccessor.Quit();
			}
		}

		private void RunInBackground_Click(object sender, RoutedEventArgs e)
		{
			if ((bool)RunInBackground.IsChecked)
				dtClockTime.Start();
			else
				dtClockTime.Stop();
		}
	}
}
