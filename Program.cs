using Microsoft.EntityFrameworkCore;
using UploadExcel.WebApi;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddDbContext<DataContext>(opt => opt.UseSqlite("Data Source=data.db"));

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

var context = app.Services.CreateScope().ServiceProvider.GetRequiredService<DataContext>();
context.Database.EnsureCreated();

app.UseHttpsRedirection();
app.MapControllers();
app.Run();
