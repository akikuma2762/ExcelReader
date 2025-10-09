using ExcelReaderAPI.Services;
using ExcelReaderAPI.Services.Interfaces;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// ✅ Phase 4: 註冊重構後的服務 (Dependency Injection)
// 核心處理服務
builder.Services.AddScoped<IExcelProcessingService, ExcelProcessingService>();
// 圖片處理服務
builder.Services.AddScoped<IExcelImageService, ExcelImageService>();
// 儲存格服務
builder.Services.AddScoped<IExcelCellService, ExcelCellService>();
// 顏色處理服務
builder.Services.AddScoped<IExcelColorService, ExcelColorService>();

// 設定CORS以允許Vue前端連接
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowVueApp", policy =>
    {
        policy.WithOrigins("http://localhost:5173", "http://localhost:5174", "http://localhost:3000") // Vue開發伺服器預設埠
              .AllowAnyMethod()
              .AllowAnyHeader()
              .AllowCredentials();
    });
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

// 啟用CORS
app.UseCors("AllowVueApp");

app.UseHttpsRedirection();
app.UseAuthorization();

app.MapControllers();

app.Run();
