using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using AlliedTestingTaskN2.Application.Excel.Commands;
using MediatR;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

namespace AlliedTestingTaskN2.Presenter.Controllers
{
    [ApiController]
    [Route("excel")]
    public class ExcelController : ControllerBase
    {
        private readonly IMediator _mediator;

        public ExcelController(IMediator mediator)
        {
            _mediator = mediator;
        }

        [HttpPost]
        [ProducesResponseType(typeof(byte[]), StatusCodes.Status200OK)]
        public async Task<ActionResult> CompareExcelFiles([FromForm] CompareExcelFilesCommand request)
        {
            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            var firstDocCheck = request.FirstExcelFile != null 
                                && request.FirstExcelFile.ContentType == contentType;
            var secondDocCheck = request.SecondExcelFile != null 
                                 && request.SecondExcelFile.ContentType == contentType;
            
            if (firstDocCheck && secondDocCheck)
            {
                var result = await _mediator.Send(request);
                return File(result, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Result");
            }
            else
            {
                return BadRequest("Both docs are required and should be in Excel format!");
            }
        }
    }
}