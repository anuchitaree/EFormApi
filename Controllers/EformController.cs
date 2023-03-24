using EFormApi.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace EFormApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EformController : ControllerBase
    {
        private readonly IExcelServices excelService;
        public EformController(IExcelServices excelService)
        {
            this.excelService = excelService;
        }


        [HttpGet("")]
        public async Task<ActionResult> GetEForm(string source = "a", string worksheet = "a")
        {
            source = @"C:\Users\Administrator\Desktop\Book10.xlsm";
            try
            {
                var result = await excelService.test(source, worksheet);
                return Ok(result);

            }
            catch (Exception)
            {

                return Ok(new { status = "ok", detail = "test" });
            }

        }

        [HttpPost("")]
        public ActionResult PostEForm([FromBody] RequestSend requestSend)
        {
            return Ok(new { status = "ok", detail = "test" });
        }

    }
}
