using TestExcel.InterfacesServices;
// set up builder with services
WebApplicationBuilder builder = WebApplication.CreateBuilder(args);
// add razor
builder.Services.AddRazorPages().AddMvcOptions(options => { });
// add interfaces to services
builder.Services.AddScoped<IExcel, SExcel>();
builder.Services.AddScoped<IExcelOpenXML, SExcelOpenXML>();
// build app
WebApplication app = builder.Build();
// configure the HTTP request pipeline in the right order
// authentication and authorization after routing and before endpoints
app.UseStatusCodePagesWithReExecute("/Error/{0}");
app.UseExceptionHandler("/Error");
app.UseHsts();
app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();
app.UseEndpoints(endpoints => { endpoints.MapRazorPages(); });
// run app
app.Run();