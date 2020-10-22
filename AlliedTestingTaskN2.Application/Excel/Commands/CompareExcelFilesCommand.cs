using System.Threading;
using System.Threading.Tasks;
using MediatR;

namespace AlliedTestingTaskN2.Application.Excel.Commands
{
    public class CompareExcelFilesCommand: IRequest<string>
    {
        public string FirstExcelFile { get; set; }
        public string SecondExcelFile { get; set; }
    }
    
    public class CompareExcelFilesCommandHandler: IRequestHandler<CompareExcelFilesCommand, string>
    {
        public async Task<string> Handle(CompareExcelFilesCommand request, CancellationToken cancellationToken)
        {
            return await Task.FromResult($"{request.FirstExcelFile} - {request.SecondExcelFile}");
        }
    }
}