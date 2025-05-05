using Microsoft.AspNetCore.Http.Features;

var builder = WebApplication.CreateBuilder(args);


builder.Services.AddControllers();

// Increase file upload size limit to 100 MB
builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 104_857_600; // 100 MB
});



var app = builder.Build();

// Redirect HTTP to HTTPS
app.UseHttpsRedirection();

app.UseAuthorization();


app.MapControllers();


app.Run();
