using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Denuo;
using SGID.Hubs;
using System.Globalization;
using SGID.Models.Inter;

var MyAllowSpecificOrigins = "_myAllowSpecificOrigins";

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
var connectionString = builder.Configuration.GetConnectionString("SGID");
var connectionProtheus = builder.Configuration.GetConnectionString("DENUO");
var connectionProtheusInter = builder.Configuration.GetConnectionString("INTER");
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(connectionString, options => options.EnableRetryOnFailure()));
builder.Services.AddDbContext<TOTVSDENUOContext>(options =>
    options.UseSqlServer(connectionProtheus, ops => 
    {
        ops.EnableRetryOnFailure();
        ops.CommandTimeout(600);
    }));
builder.Services.AddDbContext<TOTVSINTERContext>(options =>
    options.UseSqlServer(connectionProtheusInter,ops =>
    {
        ops.EnableRetryOnFailure();
        ops.CommandTimeout(600);
    }));
builder.Services.AddDatabaseDeveloperPageExceptionFilter();

builder.Services.AddIdentity<UserInter, IdentityRole>(options => options.SignIn.RequireConfirmedAccount = false)
    .AddEntityFrameworkStores<ApplicationDbContext>();

builder.Services.Configure<IdentityOptions>(options =>
{
    // Password settings.
    options.Password.RequireDigit = false;
    options.Password.RequireLowercase = false;
    options.Password.RequireNonAlphanumeric = false;
    options.Password.RequireUppercase = false;
    options.Password.RequiredLength = 6;
    options.Password.RequiredUniqueChars = 1;

    // Lockout settings.
    options.Lockout.DefaultLockoutTimeSpan = TimeSpan.FromMinutes(5);
    options.Lockout.MaxFailedAccessAttempts = 5;
    options.Lockout.AllowedForNewUsers = true;

    // User settings.
    options.User.AllowedUserNameCharacters =
    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-.@_+";
    options.User.RequireUniqueEmail = false;
    options.SignIn.RequireConfirmedEmail = false;
});

builder.Services.AddRazorPages();
builder.Services.AddSignalR();

builder.Services.AddCors(options =>
{
    options.AddPolicy(name: MyAllowSpecificOrigins,
                      policy =>
                      {
                          policy.WithOrigins("http://192.168.2.9",
                                              "https://gidd.com.br",
                                              "http://gidd.com.br",
                                              "http://200.170.251.206");
                      });
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    //app.UseExceptionHandler("/Error");
    app.UseMigrationsEndPoint();
}
else
{
    app.UseExceptionHandler("/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

var cultureInfo = new CultureInfo("en-US");
cultureInfo.NumberFormat.CurrencySymbol = "R$";

CultureInfo.DefaultThreadCurrentCulture = cultureInfo;
CultureInfo.DefaultThreadCurrentUICulture = cultureInfo;

var webSocketOptions = new WebSocketOptions
{
    KeepAliveInterval = TimeSpan.FromMinutes(5)
};

webSocketOptions.AllowedOrigins.Add("https://gidd.com.br");
webSocketOptions.AllowedOrigins.Add("http://200.170.251.206");
webSocketOptions.AllowedOrigins.Add("http://192.168.2.9");

app.UseWebSockets(webSocketOptions);

//app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

app.MapRazorPages();
app.MapHub<LogisticaHub>("/Logisticahub");

app.UseCors(MyAllowSpecificOrigins);

app.Run();