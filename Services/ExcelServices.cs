//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Spreadsheet;
using EFormApi.Interfaces;
using EFormApi.Models;
using Newtonsoft.Json;
using OfficeOpenXml;
//using ExcelDataReader;
using System.Text;
using static System.Net.Mime.MediaTypeNames;

namespace EFormApi.Services
{
    public class ExcelServices : IExcelServices
    {
        private readonly ILogger<ExcelServices> _logger;
        private readonly HttpClient _httpClient;
        public ExcelServices(ILogger<ExcelServices> logger,
            HttpClient httpClient)
        {
            _logger = logger;

            _httpClient = httpClient;

            _httpClient.BaseAddress = new Uri("https://api.github.com/");

        }
        public async Task<ReturnModel> test(string source, string worksheet)
        {

            string[] root = source.Split('\\');
            int index = root.Length - 1;

            string tempPath = "C:\\Eform";
            if (!Directory.Exists(tempPath))
            {
                Directory.CreateDirectory(tempPath);
            }

            string destFilePath = tempPath + "\\" + root[index];

            var sourceFile = new FileInfo(source);
            sourceFile.CopyTo(destFilePath, true);


            FileInfo excelFile = new FileInfo(destFilePath);
            ExcelPackage excel = new ExcelPackage(excelFile);

            //var worksheet1 = excel.Workbook.Worksheets;

            var worksheet1 = excel.Workbook.Worksheets["Sheet1"];


            if (worksheet1 == null)
            {
                return new ReturnModel
                {
                    Status = ReturnStatus.NG,
                    ReturnDetail = ReturnDetail.WrongSendData,
                    Target = source,
                };
            }
            //using (var package = new ExcelPackage(excelFile.))
            //{
            //    var worksheet = package.Workbook.Worksheets["Sheet1"];


            //var todoItem = new EFormModel()
            //{
            //    LineGroup = "",
            //    Process = "",
            //    ControlChartName = ""
            //};


            //var todoItemJson = new StringContent( JsonConvert.SerializeObject(todoItem), Encoding.UTF8,Application.Json); // using static System.Net.Mime.MediaTypeNames;

            //var httpResponseMessage =  await _httpClient.PostAsync("/api/TodoItems", todoItemJson);

            //httpResponseMessage.EnsureSuccessStatusCode();










            worksheet1.Cells["H10"].Value = worksheet1.Cells[2,5].Value;
            excel.Save();


            return new ReturnModel
            {
                Status = ReturnStatus.OK,
                ReturnDetail = ReturnDetail.Sucessfull,
                Target = source,
            };

        }
    
    }
}
