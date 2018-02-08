using System;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using RegTabWeb.Internal;
using RegTabWeb.Services;
using RegTabWeb.Utilities;

namespace RegTabWeb.Pages
{
    public class IndexModel : PageModel
    {
        private readonly IExcelFromStataRegressionLog _excelFromStataRegressionLog;
        private readonly ILogger<IndexModel> _logger;

        public const string StataLogFileFieldName = "Stata log file";

        [BindProperty]
        [Required]
        [Display(Name="Service key")]
        public string Key { get; set; }
        
        [BindProperty]
        [Required]
        [Display(Name = StataLogFileFieldName)]
        public IFormFile FileUpload { get; set; }

        [TempData]
        public string UploadedContent { get; set; }
        
        [TempData]
        public string FileName { get; set; }

        public IndexModel(
            IExcelFromStataRegressionLog excelFromStataRegressionLog,
            ILogger<IndexModel> logger)
        {
            _excelFromStataRegressionLog = excelFromStataRegressionLog;
            _logger = logger;
        }

        public IActionResult OnGet()
        {
            _logger.IndexPageRequested();
            
            if (!string.IsNullOrEmpty(UploadedContent) )
            {
                var fileDownloadName = Path.GetFileName(FileName) + "*.xlsx";

                try
                {
                    using (var stream = _excelFromStataRegressionLog.GenerateExcel(UploadedContent))
                    {
                        _logger.ExcelDownloaded(fileDownloadName);
                        return File(
                            fileContents: stream.ToArray(),
                            contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

                            // By setting a file download name the framework will
                            // automatically add the attachment Content-Disposition header
                            fileDownloadName: fileDownloadName
                        );
                    }
                }
                catch (Exception e)
                {
                    _logger.ExcelDownloadFailed(e);
                }
            }
            
            return Page();
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
            {
                return Page();
            }

            if (string.IsNullOrEmpty(FileUpload?.FileName))
            {
                return Page();
            }
            
            if (Key != "cindy")
            {
                return Page();
            }
            
            _logger.StataLogFileUploadRequested(FileUpload.FileName);

            try
            {
                FileName = FileUpload.FileName;
                
                UploadedContent = await FileHelpers.ProcessFormFile(FileUpload, ModelState);
            }
            catch (Exception e)
            {
                _logger.LogFileUploadFailed(e);
            }

            return RedirectToPage();
        }
    }
}