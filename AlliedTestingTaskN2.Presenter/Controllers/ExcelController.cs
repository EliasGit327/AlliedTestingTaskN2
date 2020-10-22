using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using AlliedTestingTaskN2.Application.Excel.Commands;
using MediatR;
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
        public async Task<ActionResult<string>> CompareExcelFiles([FromBody] CompareExcelFilesCommand request)
        {
            var result = await _mediator.Send(request);
            return result != null ? Ok(result) : BadRequest() as ActionResult;
        }
    }
}