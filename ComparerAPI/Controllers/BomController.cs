using Microsoft.AspNetCore.Mvc;

namespace ComparerAPI
{
    [ApiController]
    [Route("api/[controller]")]
    public class BomController : ControllerBase
    {
        private readonly IBomFileService _bomFileService;

        public BomController(IBomFileService bomFileService)
        {
            _bomFileService = bomFileService;
        }

        [HttpPost("ReadAndCompare")]
        public ActionResult CompareBomFiles(string sourcePath, string targetPath)
        {
            try
            {
                var comparedBomFile = _bomFileService.CompareBomFiles(sourcePath, targetPath);
                var excelBytes = _bomFileService.WriteComparedBomToExcel(comparedBomFile);

                return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "report.xlsx");
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, $"An error occurred: {ex.Message}");
            }
        }
    }
}