using BOMComparer;
using BOMComparer.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Org.BouncyCastle.Utilities;
using System.IO;

namespace ComparerAPI.Controllers
{


    [Route("api/[controller]")]
    [ApiController]
    public class BomController : ControllerBase
    {      

        [HttpPost("ReadBomFile")]
        public ActionResult<BomFile> ReadBomFile( string filePath, string filePath2)
        {
            ExcelReader excelReader = new();
            BomFile bomFile = excelReader.ReadBomFile(filePath);
            BomFile bomFile2 = excelReader.ReadBomFile(filePath2);
            Comparer comparer = new();
            var comparedBom = comparer.ComparedBomFile(bomFile, bomFile2);
            ExcelWriter excelWriter = new();
            var b = excelWriter.WriteExcelFile(comparedBom); 
            
            return File( b, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",  "report.xlsx");            

        }
    }
}