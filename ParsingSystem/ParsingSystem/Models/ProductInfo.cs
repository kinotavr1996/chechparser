using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParsingSystem.Models
{
	public class ProductInfo
	{
		public decimal YourPrice { get; set; }
		public decimal LowestPrice { get; set; }
		public List<decimal> Prices { get; set; }
		public int ItemId { get; set; }
		public string Url { get; set; }
	}
}
