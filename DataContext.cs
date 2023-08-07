using Microsoft.EntityFrameworkCore;

namespace UploadExcel.WebApi;

public class DataContext : DbContext
{
	public DataContext(DbContextOptions<DataContext> opt) : base(opt)
	{
	}

    public DbSet<Product> Products { get; set; }
}
