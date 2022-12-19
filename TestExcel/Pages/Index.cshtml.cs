using Microsoft.AspNetCore.Mvc.RazorPages;
using TestExcel.InterfacesServices;
using TestExcel.Models.Views;

namespace TestExcel.Pages
{
    public class IndexModel : PageModel
    {
        public string Mess { get; set; } = string.Empty;
        public List<TestView> ListRows { get; set; } = new();

        private readonly IExcel _iExcel;
        private readonly IHostEnvironment _iHostEnvironment;
        public IndexModel(IExcel iExcel, IHostEnvironment iHostEnvironment)
        {
            _iExcel = iExcel;
            _iHostEnvironment = iHostEnvironment;
        }
        public void OnGet()
        {
        }
        public void OnPostClear()
        {
            // clear message and data rows
            Mess = string.Empty;
            ListRows = new();
        }
        public void OnPostRead()
        {
            // set file name and path
            string fnap = Path.Combine(_iHostEnvironment.ContentRootPath, "test.xlsx");
            // read spreadsheet with this file name and path
            Tuple<string, List<TestView>> results = _iExcel.Read(fnap);
            // set message and data rows from results
            Mess = results.Item1;
            ListRows = results.Item2;
        }
        public void OnPostWrite()
        {
            // set file name and path
            string fnap = Path.Combine(_iHostEnvironment.ContentRootPath, "test.xlsx");
            // set data rows
            ListRows = new()
            {
                new()
                {
                    ABool = true,
                    ADateTimeNullable = DateTime.Now,
                    ADecimalNullable = 123.456m,
                    ADouble = 345.67,
                    AnInt = 9,
                    ANumberAsText = "000123",
                    AString = "Hello world"
                },
                new()
                {
                    ABool = false,
                    ADateTimeNullable = DateTime.Now.AddDays(1),
                    ADecimalNullable = 654.321m,
                    ADouble = 765.43,
                    AnInt = 7,
                    ANumberAsText = "000456",
                    AString = "Goodbye cruel world"
                }
            };
            // write spreadsheet with this file name and path and these data rows
            Mess = _iExcel.Write(ListRows, fnap);
        }
    }
}