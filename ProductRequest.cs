namespace UploadExcel.WebApi;

public class ProductRequest
{
    public string Name { get; set; }
    public int Quantity { get; set; }
    public decimal Price { get; set; }
    public bool IsActive { get; set; }
    public DateTime ExpiryDate { get; set; }
}
