using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Diagnostics;

namespace TestExcel.Pages
{
    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    [IgnoreAntiforgeryToken]
    public class ErrorModel : PageModel
    {
        public string? RequestId { get; set; }

        public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);

        public string Mess { get; set; } = string.Empty;
        public void OnGet(int? id)
        {
            Mess = id switch
            {
                401 => "Your credentials could not be confirmed. Please try again later, or raise a ticket in TopDesk.",
                403 => "You do not have permission to do this.",
                404 => "The page requested is wrong.",
                408 => "The page took too long to load. Please try again later, or raise a ticket in TopDesk.",
                _ => "Something has gone wrong. Please try again later, or raise a ticket in TopDesk.",
            };
        }
    }
}