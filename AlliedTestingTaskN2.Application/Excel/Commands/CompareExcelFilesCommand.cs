using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using MediatR;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace AlliedTestingTaskN2.Application.Excel.Commands
{
    public class CompareExcelFilesCommand : IRequest<byte[]>
    {
        public IFormFile FirstExcelFile { get; set; }
        public IFormFile SecondExcelFile { get; set; }
    }

    public class CompareExcelFilesCommandHandler : IRequestHandler<CompareExcelFilesCommand, byte[]>
    {
        private readonly IHostingEnvironment _hostingEnvironment;

        public CompareExcelFilesCommandHandler(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        public async Task<byte[]> Handle(CompareExcelFilesCommand request, CancellationToken cancellationToken)
        {
            ExcelPackage excelFirstFile;
            ExcelPackage excelSecondFile;

            await using (var memoryStream = new MemoryStream())
            {
                await request.FirstExcelFile.CopyToAsync(memoryStream, cancellationToken);
                excelFirstFile = new ExcelPackage(memoryStream);
            }
            var workSheetOfFirstFile = excelFirstFile.Workbook.Worksheets[0];
            
            await using (var memoryStream = new MemoryStream())
            {
                await request.SecondExcelFile.CopyToAsync(memoryStream, cancellationToken);
                excelSecondFile = new ExcelPackage(memoryStream);
            }
            var workSheetOfSecondFile = excelSecondFile.Workbook.Worksheets[0];

            var start = workSheetOfSecondFile.Dimension.Start;
            var end = workSheetOfSecondFile.Dimension.End;
            for (var row = start.Row; row <= end.Row; row++)
            {
                for (var col = start.Column; col <= end.Column; col++)
                {
                    var firstValue = workSheetOfFirstFile.Cells[row, col].Value;
                    var secondValue = workSheetOfSecondFile.Cells[row, col].Value;

                    if (firstValue == null && secondValue != null)
                    {
                        workSheetOfFirstFile.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheetOfFirstFile.Cells[row, col].Style.Fill.BackgroundColor
                            .SetColor(ColorTranslator.FromHtml("red"));
                        workSheetOfFirstFile.Cells[row, col].Value = $"none - {secondValue}";
                    }
                    
                    if (firstValue != null && secondValue == null)
                    {
                        workSheetOfFirstFile.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheetOfFirstFile.Cells[row, col].Style.Fill.BackgroundColor
                            .SetColor(ColorTranslator.FromHtml("red"));
                        workSheetOfFirstFile.Cells[row, col].Value = $"{firstValue} - none";
                    }
                    
                    if (firstValue != null && secondValue != null)
                    {
                        if (firstValue.ToString() != secondValue.ToString())
                        {
                            workSheetOfFirstFile.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workSheetOfFirstFile.Cells[row, col].Style.Fill.BackgroundColor
                                .SetColor(ColorTranslator.FromHtml("red"));
                            workSheetOfFirstFile.Cells[row, col].Value = $"{firstValue} - {secondValue}";
                        }
                    }
                }
            }


            var bytesOfFirstFile = await excelFirstFile.GetAsByteArrayAsync(cancellationToken);
            var bytesOfSecondFile = await excelSecondFile.GetAsByteArrayAsync(cancellationToken);


            return await Task.FromResult(bytesOfFirstFile);
        }
    }
}