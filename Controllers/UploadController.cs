using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.IO;
using System.Threading.Tasks;

namespace Backend.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class UploadController : ControllerBase
    {
        [HttpPost("upload")]
        public async Task<IActionResult> UploadFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file was uploaded.");
            }

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using var inputStream = new MemoryStream();
                await file.CopyToAsync(inputStream);
                inputStream.Position = 0;

                using var package = new ExcelPackage(inputStream);

                if (package.Workbook.Worksheets.Count == 0)
                {
                    return BadRequest("The uploaded Excel file contains no worksheets.");
                }

                var worksheet = package.Workbook.Worksheets[0]; 
                var rowCount = worksheet.Dimension?.Rows ?? 0;
                var colCount = worksheet.Dimension?.Columns ?? 0;

                if (rowCount == 0 || colCount < 2) 
                {
                    return BadRequest("The Excel file must contain at least two columns: 'Num1' and 'Num2'.");
                }

                using var outputPackage = new ExcelPackage();
                var outputSheet = outputPackage.Workbook.Worksheets.Add("Output");

               
                outputSheet.Cells[1, 1].Value = "Num1";
                outputSheet.Cells[1, 2].Value = "Num2";
                outputSheet.Cells[1, 3].Value = "Output";

               
                for (int row = 2; row <= rowCount; row++)
                {
                    var num1Text = worksheet.Cells[row, 1]?.Text;
                    var num2Text = worksheet.Cells[row, 2]?.Text;

                    int.TryParse(num1Text, out int num1);
                    int.TryParse(num2Text, out int num2);

                   
                    outputSheet.Cells[row, 1].Value = num1;
                    outputSheet.Cells[row, 2].Value = num2;
                    outputSheet.Cells[row, 3].Value = num1 + num2; 
                }

                var fileStream = new MemoryStream();
                outputPackage.SaveAs(fileStream);
                fileStream.Position = 0;

                return File(
                    fileStream,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "output.xlsx"
                );
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"An error occurred: {ex.Message}");
            }
        }
    }
}
