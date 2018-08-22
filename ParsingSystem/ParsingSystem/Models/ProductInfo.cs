namespace ParsingSystem.Models
{
	public class ProductInfo
	{
		public decimal Price { get; set; }
		public decimal PostageCost { get; set; }
		public decimal LowestPrice { get; set; }
		public int ItemId { get; set; }
		public string Url { get; set; }
		public string Category { get; set; }
		public string Description { get; set; }
		public bool IsParsed { get; set; }
	}
}
