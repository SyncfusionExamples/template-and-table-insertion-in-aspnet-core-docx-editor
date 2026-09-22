using Newtonsoft.Json;
using Syncfusion.EJ2.SpellChecker;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.AspNetCore.Rewrite;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();


// Increase upload/request size limit to 500 MB
builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 500 * 1024 * 1024;
});

builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = 500 * 1024 * 1024;
});


var configuration = builder.Configuration;
var env = builder.Environment;

// Load spell check configuration
string path = configuration["SPELLCHECK_DICTIONARY_PATH"];
string jsonFileName = configuration["SPELLCHECK_JSON_FILENAME"];

// Set default path if not provided
path = string.IsNullOrEmpty(path) ? Path.Combine(env.ContentRootPath, "App_Data") : Path.Combine(env.ContentRootPath, path);
jsonFileName = string.IsNullOrEmpty(jsonFileName) ? Path.Combine(path, "spellcheck.json") : Path.Combine(path, jsonFileName);

if (File.Exists(jsonFileName))
{
    string jsonImport = File.ReadAllText(jsonFileName);
    List<DictionaryData> spellChecks = JsonConvert.DeserializeObject<List<DictionaryData>>(jsonImport);
    List<DictionaryData> spellDictCollection = new List<DictionaryData>();
    string personalDictPath = null;

    if (spellChecks != null)
    {
        foreach (var spellCheck in spellChecks)
        {
            spellDictCollection.Add(new DictionaryData(
                spellCheck.LanguadeID,
                Path.Combine(path, spellCheck.DictionaryPath),
                Path.Combine(path, spellCheck.AffixPath)
            ));
            personalDictPath = Path.Combine(path, spellCheck.PersonalDictPath);
        }
    }

    // SpellChecker.InitializeDictionaries(spellDictCollection, personalDictPath, 3);
}

var app = builder.Build();

Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("");

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();

app.UseDefaultFiles();
app.UseStaticFiles();
app.MapControllers();


var basePath = "/template-and-table-insertion-in-aspnet-core-docx-editor";
app.UsePathBase(basePath);

app.UseStaticFiles(new StaticFileOptions
{
    ServeUnknownFileTypes = true
});

app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
