using ClosedXML.Excel;
using Integradas.Dtos;
using Integradas.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace Integradas.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class IntegratedController : ControllerBase
    {
        private readonly AppDbContext _context;

        public IntegratedController(AppDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        [Route("Weeks/summary")]
        public async Task<IActionResult> GetWeeksSummary()
        {
            try
            {
                var weeks = await _context.Integradas
                    .Where(x => x.WeekNumber.HasValue)
                    .GroupBy(x => x.WeekNumber!.Value)
                    .Select(g => new OrderSummaryDto
                    {
                        WeekNumber = g.Key,
                        TotalOrders = g.Count(),
                        CompletedOrders = g.Count(x => x.Completed),
                        PendingOrders = g.Count(x => !x.Completed),
                        TotalQuantity = g.Sum(x => x.Amount ?? 0),
                        ScannedQuantity = g.Sum(x => x.ScannedQuantity ?? 0),
                        ProgressPercentage = g.Sum(x => x.Amount ?? 0) > 0 ?
                            (double)g.Sum(x => x.ScannedQuantity ?? 0) / g.Sum(x => x.Amount ?? 0) * 100 : 0,
                        LastUpdate = g.Max(x => x.CompletedDate ?? x.CreatedAt)
                    })
                    .OrderByDescending(x => x.WeekNumber)
                    .ToListAsync();

                return Ok(new ApiResponse<List<OrderSummaryDto>>
                {
                    Success = true,
                    Message = "Resumen de semanas obtenido exitosamente",
                    Data = weeks
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new ApiResponse<List<OrderSummaryDto>>
                {
                    Success = false,
                    Message = $"Error al obtener el resumen: {ex.Message}"
                });
            }
        }

        [HttpGet]
        [Route("Week/{weekNumber}")]
        public async Task<IActionResult> GetWeekDetails(int weekNumber)
        {
            try
            {
                var orders = await _context.Integradas
                    .Where(x => x.WeekNumber == weekNumber)
                    .OrderBy(x => x.PartNumber)
                    .Select(x => new OrderDetailDto
                    {
                        Id = x.Id,
                        Type = x.Type ?? string.Empty,
                        PartNumber = x.PartNumber ?? string.Empty,
                        Pipeline = x.Pipeline ?? string.Empty,
                        Amount = x.Amount ?? 0,
                        ScannedQuantity = x.ScannedQuantity ?? 0,
                        Completed = x.Completed,
                        CreatedAt = x.CreatedAt,
                        CompletedDate = x.CompletedDate
                    })
                    .ToListAsync();

                if (!orders.Any())
                {
                    return NotFound(new ApiResponse<List<OrderDetailDto>>
                    {
                        Success = false,
                        Message = $"No se encontraron órdenes para la semana {weekNumber}"
                    });
                }

                return Ok(new ApiResponse<List<OrderDetailDto>>
                {
                    Success = true,
                    Message = $"Órdenes de la semana {weekNumber} obtenidas exitosamente",
                    Data = orders
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new ApiResponse<List<OrderDetailDto>>
                {
                    Success = false,
                    Message = $"Error al obtener las órdenes: {ex.Message}"
                });
            }
        }

        [HttpGet]
        [Route("GenerateReport/{weekNumber}")]
        public async Task<IActionResult> GenerateReport(int weekNumber, [FromQuery] string reportType = "detailed")
        {
            try
            {
                if(reportType != "detailed" && reportType != "summary")
                {
                    return BadRequest(new ReportResponseDto
                    {
                        Success = false,
                        Message = "Tipo de reporte no válido. Use detailed o summary"
                    });
                }

                var weekData = await _context.Integradas
                            .Where(x => x.WeekNumber == weekNumber)
                            .OrderBy(x => x.PartNumber)
                            .ToListAsync();

                if(!weekData.Any())
                {
                    return NotFound(new ReportResponseDto
                    {
                        Success = false,
                        Message = $"No se encontraron datos para la semana {weekNumber}"
                    });
                }

                using (var workbook = new XLWorkbook())
                {
                    if(reportType == "detailed")
                    {
                        GenerateDetailedReport(workbook, weekData, weekNumber);
                    }
                    else
                    {
                        GenerateSummaryReport(workbook, weekData, weekNumber);
                    }

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();

                        var fileName = $"Reporte_Semana_{weekNumber}_{DateTime.Now:ddMMyyyy_HHmmss}.xlsx";

                        return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    }

                }
            }
            catch(Exception ex)
            {
                return StatusCode(500, new ReportResponseDto
                {
                    Success = false,
                    Message = $"Error generando el reporte: {ex.Message}"
                });
            }
        }

        [HttpPost]
        [Route("Scan")]
        public async Task<IActionResult> ScanPart([FromBody] ScanRequestDto request)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.PartNumber))
                {
                    return BadRequest(new ScanResponseDto
                    {
                        Success = false,
                        Message = "El número de parte es requerido"
                    });
                }

                if (request.Quantity <= 0)
                {
                    return BadRequest(new ScanResponseDto
                    {
                        Success = false,
                        Message = "La cantidad debe ser mayor a 0"
                    });
                }

                var order = await _context.Integradas
                    .FirstOrDefaultAsync(o =>
                        o.PartNumber == request.PartNumber.Trim() &&
                        o.WeekNumber == request.WeekNumber);

                if (order == null)
                {
                    return NotFound(new ScanResponseDto
                    {
                        Success = false,
                        Message = $"No se encontró la orden para el número de parte {request.PartNumber} en la semana {request.WeekNumber}"
                    });
                }

                if (order.Completed)
                {
                    return BadRequest(new ScanResponseDto
                    {
                        Success = false,
                        Message = $"La orden {order.PartNumber} ya está completada"
                    });
                }

                var currentScanned = order.ScannedQuantity ?? 0;
                var requiredAmount = order.Amount ?? 0;

                var newScannedQuantity = currentScanned + request.Quantity;

                if (newScannedQuantity > requiredAmount)
                {
                    var excess = newScannedQuantity - requiredAmount;
                    return BadRequest(new ScanResponseDto
                    {
                        Success = false,
                        Message = $"No se pueden agregar {request.Quantity} unidades. Excede la cantidad requerida por {excess} unidades. Máximo permitido: {requiredAmount - currentScanned}"
                    });
                }

                order.ScannedQuantity = newScannedQuantity;

                if (order.ScannedQuantity >= requiredAmount)
                {
                    order.Completed = true;
                    order.CompletedDate = DateTime.UtcNow;
                }

                await _context.SaveChangesAsync();

                var updatedScanned = order.ScannedQuantity ?? 0;
                var remaining = Math.Max(0, requiredAmount - updatedScanned);
                var progressPercentage = requiredAmount > 0 ?
                    (double)updatedScanned / requiredAmount * 100 : 0;

                var orderDto = new OrderDetailDto
                {
                    Id = order.Id,
                    Type = order.Type ?? string.Empty,
                    PartNumber = order.PartNumber ?? string.Empty,
                    Pipeline = order.Pipeline ?? string.Empty,
                    Amount = requiredAmount,
                    ScannedQuantity = updatedScanned,
                    Completed = order.Completed,
                    CreatedAt = order.CreatedAt,
                    CompletedDate = order.CompletedDate
                };

                return Ok(new ScanResponseDto
                {
                    Success = true,
                    Message = order.Completed ?
                        $"¡COMPLETADO! {order.PartNumber} - Se agregaron {request.Quantity} unidades. Total: {updatedScanned}/{requiredAmount}" :
                        $"Escaneado: Se agregaron {request.Quantity} unidades. Total: {updatedScanned}/{requiredAmount}",
                    PartNumber = order.PartNumber ?? string.Empty,
                    ScannedQuantity = updatedScanned,
                    RequiredQuantity = requiredAmount,
                    AddedQuantity = request.Quantity,
                    Completed = order.Completed,
                    Remaining = remaining,
                    ProgressPercentage = progressPercentage,
                    Order = orderDto
                });
            }
            catch (Exception ex)
            {
            
                return StatusCode(500, new ScanResponseDto
                {
                    Success = false,
                    Message = $"Error al escanear: {ex.Message}"
                });
            }
        }

        [HttpPost]
        [Route("ProductionOrderIncrease")]
        public async Task<IActionResult> ProductionOrderIncrease([FromForm] ExcelUploadRequest request)
        {
            var response = new ProductionOrderResponse();

            try
            {
                if(request.File == null || request.File.Length == 0)
                {
                    response.Success = false;
                    response.Message = "No se ha porpocionado ningun archivo";
                    return BadRequest(response);
                }

                var allowedExtension = new[] { ".xlsx", ".xls" };
                var fileExtension = Path.GetExtension(request.File.FileName).ToLower();

                if(!allowedExtension.Contains(fileExtension))
                {
                    response.Success = false;
                    response.Message = "El archivo debe ser un Excel (.xlsx o .xls)";
                    return BadRequest(response);
                }

                using (var stream = new MemoryStream())
                {
                    await request.File.CopyToAsync(stream);

                    using (var workbook = new XLWorkbook(stream))
                    {
                        var worksheet = workbook.Worksheet(1);

                        if (worksheet.RangeUsed() == null)
                        {
                            response.Success = false;
                            response.Message = "El archivo Excel está vacío";
                            return BadRequest(response);
                        }

                        var records = new List<IntegradasDto>();
                        var rowCount = worksheet.RangeUsed()!.RowCount();

                        for (int row = 2; row <= rowCount; row++)
                        {
                            try
                            {
                                var typeCell = worksheet.Cell(row, 1);
                                var partNumberCell = worksheet.Cell(row, 2);
                                var pipelineCell = worksheet.Cell(row, 3);
                                var amountCell = worksheet.Cell(row, 4);

                                var record = new IntegradasDto
                                {
                                    Type = typeCell.GetValue<string>()?.Trim()!,
                                    PartNumber = partNumberCell.GetValue<string>()?.Trim()!,
                                    Pipeline = pipelineCell.GetValue<string>()?.Trim()!,
                                    Amount = amountCell.GetValue<int?>() ?? ParseInt(amountCell.GetValue<string>()),
                                    WeekNumber = request.WeekNumber
                                };

                                if (!string.IsNullOrWhiteSpace(record.PartNumber) && record.Amount.HasValue)
                                {
                                    records.Add(record);
                                }
                                else
                                {
                                    response.Errors.Add($"Fila {row}: Datos incompletos (PART NBR. y Cantidad son requeridos)");
                                }
                            }
                            catch (Exception ex)
                            {
                                response.Errors.Add($"Error en fila {row}: {ex.Message}");
                            }
                        }

                        if (records.Any())
                        {
                            await SaveRecordsToDatabase(records);
                            response.RecordProcessed = records.Count;
                            response.Success = true;
                            response.Message = $"{records.Count} registros procesados exitosamente";

                            if (response.Errors.Any())
                            {
                                response.Message += $". Se encontraron {response.Errors.Count} advertencias";
                            }
                        }
                        else
                        {
                            response.Success = false;
                            response.Message = "No se encontraron registros válidos en el archivo";
                            return BadRequest(response);
                        }
                    }
                }

                return Ok(response);
            }
            catch (Exception ex)
            {
                response.Success = false;
                response.Message = $"Error procesando el archivo: {ex.Message}";
                return StatusCode(500, response);
            }
        }

        private int? ParseInt(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;

            value = value.Trim();

            if (int.TryParse(value, out int result))
            {
                return result;
            }

            if (double.TryParse(value, out double doubleValue))
            {
                return (int)Math.Round(doubleValue);
            }

            return null;
        }

        private async Task SaveRecordsToDatabase(List<IntegradasDto> records)
        {
            try
            {
                var entities = records.Select(record => new Models.Integradas
                {
                    Type = record.Type,
                    PartNumber = record.PartNumber,
                    Pipeline = record.Pipeline,
                    Amount = record.Amount!.Value,
                    WeekNumber = record.WeekNumber,
                }).ToList();

                await _context.Integradas.AddRangeAsync(entities);
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateException dbEx)
            {
                throw new Exception($"Error de base de datos: {dbEx.InnerException?.Message ?? dbEx.Message}", dbEx);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error guardando registros: {ex.Message}", ex);
            }
        }

        private void GenerateDetailedReport(XLWorkbook workbook, List<Models.Integradas> weekData, int weekNumber)
        {
            var worksheet = workbook.Worksheets.Add($"Semana {weekNumber} - Detallado");

            var titleStyle = workbook.Style;
            titleStyle.Font.Bold = true;
            titleStyle.Font.FontSize = 16;
            titleStyle.Font.FontColor = XLColor.White;
            titleStyle.Fill.BackgroundColor = XLColor.FromHtml("#2E86AB");
            titleStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titleStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            var subtitleStyle = workbook.Style;
            subtitleStyle.Font.Bold = true;
            subtitleStyle.Font.FontSize = 12;
            subtitleStyle.Font.FontColor = XLColor.FromHtml("#2E86AB");
            subtitleStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

            var headerStyle = workbook.Style;
            headerStyle.Font.Bold = true;
            headerStyle.Font.FontSize = 10;
            headerStyle.Font.FontColor = XLColor.White;
            headerStyle.Fill.BackgroundColor = XLColor.FromHtml("#A23B72");
            headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerStyle.Border.OutsideBorderColor = XLColor.Black;

            var dataStyle = workbook.Style;
            dataStyle.Font.FontSize = 9;
            dataStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            dataStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;
            dataStyle.Border.OutsideBorderColor = XLColor.Gray;
            
            var alternateStyle = workbook.Style;
            alternateStyle.Font.FontSize = 9;
            alternateStyle.Fill.BackgroundColor = XLColor.FromHtml("#F8F9FA");
            alternateStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            alternateStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;
            alternateStyle.Border.OutsideBorderColor= XLColor.Gray;

            var completedStyle = workbook.Style;
            completedStyle.Font.FontSize = 9;
            completedStyle.Font.FontColor = XLColor.FromHtml("#28A745");
            completedStyle.Fill.BackgroundColor = XLColor.FromHtml("#D4EDDA");
            completedStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            completedStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;
            completedStyle.Border.OutsideBorderColor = XLColor.Gray;

            var pendingStyle = workbook.Style;
            pendingStyle.Font.FontSize = 9;
            pendingStyle.Font.FontColor = XLColor.FromHtml("#DC3545");
            pendingStyle.Fill.BackgroundColor = XLColor.FromHtml("#F8D7DA");
            pendingStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            pendingStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;
            pendingStyle.Border.OutsideBorderColor = XLColor.Gray;

            worksheet.Cell("A1").Value = $"REPORTE DE PRODUCCIÓN - SEMANA {weekNumber}";
            worksheet.Range("A1:L1").Merge();
            worksheet.Range("A1:L1").Style = titleStyle;
            worksheet.Row(1).Height = 25;

            worksheet.Cell("A2").Value = "Integradas L-18";
            worksheet.Cell("A2").Style = subtitleStyle;
            worksheet.Cell("E2").Value = "Fecha de generación:";
            worksheet.Cell("F2").Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
            worksheet.Cell("F2").Style.Font.Bold = true;

            worksheet.Cell("A3").Value = "Sistema de Gestión de Producción";
            worksheet.Cell("A3").Style.Font.Italic = true;
            worksheet.Cell("E3").Value = "Total de órdenes:";
            worksheet.Cell("F3").Value = weekData.Count;
            worksheet.Cell("F3").Style.Font.Bold = true;

            var completedCount = weekData.Count(x => x.Completed);
            var pendingCount = weekData.Count(x => !x.Completed);
            var totalRequired = weekData.Sum(x => x.Amount ?? 0);
            var totalScanned = weekData.Sum(x => x.ScannedQuantity ?? 0);
            var progressPercentage = totalRequired > 0 ? (double)totalScanned / totalRequired * 100 : 0;

            worksheet.Cell("I2").Value = "Completadas:";
            worksheet.Cell("J2").Value = completedCount;
            worksheet.Cell("J2").Style.Font.Bold = true;
            worksheet.Cell("J2").Style.Fill.BackgroundColor = XLColor.FromHtml("#D4EDDA");

            worksheet.Cell("I3").Value = "Pendientes:";
            worksheet.Cell("J3").Value = pendingCount;
            worksheet.Cell("J3").Style.Font.Bold = true;
            worksheet.Cell("J3").Style.Fill.BackgroundColor = XLColor.FromHtml("#F8D7DA");

            worksheet.Cell("K2").Value = "Progreso total:";
            worksheet.Cell("L2").Value = $"{progressPercentage:F1}%";
            worksheet.Cell("L2").Style.Font.Bold = true;

            worksheet.Cell("K3").Value = "Total escaneado:";
            worksheet.Cell("L3").Value = $"{totalScanned}/{totalRequired}";
            worksheet.Cell("L3").Style.Font.Bold = true;

            var headers = new[]
            {
                "ID", "Número de Parte", "Tipo", "Tubería (OD Pulg)",
                "Cantidad Requerida", "Cantidad Escaneada", "Restante",
                "Progreso %", "Estado", "Fecha Creación", "Fecha Completado", "Semana"
            };

            int currentRow = 5;
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cell(currentRow, i + 1).Value = headers[i];
                worksheet.Cell(currentRow, i + 1).Style = headerStyle;
                worksheet.Column(i + 1).Width = 15;
            }

            worksheet.Column(2).Width = 20;
            worksheet.Column(4).Width = 18;
            worksheet.Column(10).Width = 18;
            worksheet.Column(11).Width = 18;

            currentRow++;
            foreach (var order in weekData)
            {
                var row = worksheet.Row(currentRow);
                row.Height = 20;

                var progress = order.Amount > 0 ? (double)(order.ScannedQuantity ?? 0) / order.Amount * 100 : 0;
                var remaining = Math.Max(0, (order.Amount ?? 0) - (order.ScannedQuantity ?? 0));
                var status = order.Completed ? "COMPLETADO" : "EN PROGRESO";

                worksheet.Cell(currentRow, 1).Value = order.Id;
                worksheet.Cell(currentRow, 2).Value = order.PartNumber;
                worksheet.Cell(currentRow, 3).Value = order.Type;
                worksheet.Cell(currentRow, 4).Value = order.Pipeline;
                worksheet.Cell(currentRow, 5).Value = order.Amount;
                worksheet.Cell(currentRow, 6).Value = order.ScannedQuantity;
                worksheet.Cell(currentRow, 7).Value = remaining;
                worksheet.Cell(currentRow, 8).Value = progress;
                worksheet.Cell(currentRow, 9).Value = status;
                worksheet.Cell(currentRow, 10).Value = order.CreatedAt?.ToString("dd/MM/yyyy HH:mm");
                worksheet.Cell(currentRow, 11).Value = order.CompletedDate?.ToString("dd/MM/yyyy HH:mm");
                worksheet.Cell(currentRow, 12).Value = order.WeekNumber;

                IXLStyle cellStyle;

                if (order.Completed)
                {
                    cellStyle = completedStyle;
                }
                else if (remaining == 0 && (order.ScannedQuantity ?? 0) >= (order.Amount ?? 0))
                {
                    cellStyle = completedStyle;
                }
                else
                {
                    cellStyle = (currentRow % 2 == 0) ? dataStyle : alternateStyle;
                }

                for (int col = 1; col <= 12; col++)
                {
                    worksheet.Cell(currentRow, col).Style = cellStyle;

                    if (col == 8)
                    {
                        worksheet.Cell(currentRow, col).Style.NumberFormat.Format = "0.0%";
                    }

                    if (col >= 5 && col <= 8) // Columnas numéricas
                    {
                        worksheet.Cell(currentRow, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }
                }

                currentRow++;
            }

            currentRow += 2;
            worksheet.Cell(currentRow, 1).Value = "RESUMEN EJECUTIVO";
            worksheet.Range($"A{currentRow}:L{currentRow}").Merge();
            worksheet.Range($"A{currentRow}:L{currentRow}").Style = headerStyle;

            currentRow++;
            worksheet.Cell(currentRow, 1).Value = "Métrica";
            worksheet.Cell(currentRow, 2).Value = "Valor";
            worksheet.Cell(currentRow, 1).Style = headerStyle;
            worksheet.Cell(currentRow, 2).Style = headerStyle;

            var metrics = new[]
            {
                ("Total de Órdenes", weekData.Count.ToString()),
                ("Órdenes Completadas", $"{completedCount} ({((double)completedCount/weekData.Count*100):F1}%)"),
                ("Órdenes Pendientes", $"{pendingCount} ({((double)pendingCount/weekData.Count*100):F1}%)"),
                ("Cantidad Total Requerida", totalRequired.ToString()),
                ("Cantidad Total Escaneada", totalScanned.ToString()),
                ("Progreso General", $"{progressPercentage:F1}%"),
                ("Eficiencia de Producción", $"{((double)completedCount/weekData.Count*100):F1}%")
            };

            foreach (var (metric, value) in metrics)
            {
                currentRow++;
                worksheet.Cell(currentRow, 1).Value = metric;
                worksheet.Cell(currentRow, 2).Value = value;
                worksheet.Cell(currentRow, 1).Style = dataStyle;
                worksheet.Cell(currentRow, 2).Style = dataStyle;
                worksheet.Cell(currentRow, 2).Style.Font.Bold = true;
            }

            currentRow += 2;
            worksheet.Cell(currentRow, 1).Value = "Generado por:";
            worksheet.Cell(currentRow, 6).Value = "Revisado por:";
            worksheet.Cell(currentRow, 11).Value = "Aprobado por:";

            currentRow += 3;
            worksheet.Cell(currentRow, 1).Value = "_________________________";
            worksheet.Cell(currentRow, 6).Value = "_________________________";
            worksheet.Cell(currentRow, 11).Value = "_________________________";

            worksheet.SheetView.FreezeRows(5);
        }

        private void GenerateSummaryReport(XLWorkbook workbook, List<Models.Integradas> weekData, int weekNumber)
        {
            var worksheet = workbook.Worksheets.Add($"Semana {weekNumber} - Resumen");

            var titleStyle = workbook.Style;
            titleStyle.Font.Bold = true;
            titleStyle.Font.FontSize = 14;
            titleStyle.Font.FontColor = XLColor.White;
            titleStyle.Fill.BackgroundColor = XLColor.FromHtml("#2E86AB");
            titleStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            var headerStyle = workbook.Style;
            headerStyle.Font.Bold = true;
            headerStyle.Font.FontSize = 10;
            headerStyle.Font.FontColor = XLColor.White;
            headerStyle.Fill.BackgroundColor = XLColor.FromHtml("#A23B72");
            headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            worksheet.Cell("A1").Value = $"REPORTE RESUMEN - SEMANA {weekNumber}";
            worksheet.Range("A1:F1").Merge();
            worksheet.Range("A1:F1").Style = titleStyle;

            var statsByType = weekData
                .GroupBy(x => x.Type ?? "Sin Tipo")
                .Select(g => new
                {
                    Type = g.Key,
                    TotalOrders = g.Count(),
                    Completed = g.Count(x => x.Completed),
                    TotalRequired = g.Sum(x => x.Amount ?? 0),
                    TotalScanned = g.Sum(x => x.ScannedQuantity ?? 0),
                    Progress = g.Sum(x => x.Amount ?? 0) > 0 ?
                        (double)g.Sum(x => x.ScannedQuantity ?? 0) / g.Sum(x => x.Amount ?? 0) * 100 : 0
                })
                .OrderByDescending(x => x.TotalOrders)
                .ToList();

            worksheet.Cell("A3").Value = "Tipo";
            worksheet.Cell("B3").Value = "Total Órdenes";
            worksheet.Cell("C3").Value = "Completadas";
            worksheet.Cell("D3").Value = "Cantidad Requerida";
            worksheet.Cell("E3").Value = "Cantidad Escaneada";
            worksheet.Cell("F3").Value = "Progreso %";

            worksheet.Range("A3:F3").Style = headerStyle;

            int row = 4;
            foreach (var stat in statsByType)
            {
                worksheet.Cell(row, 1).Value = stat.Type;
                worksheet.Cell(row, 2).Value = stat.TotalOrders;
                worksheet.Cell(row, 3).Value = stat.Completed;
                worksheet.Cell(row, 4).Value = stat.TotalRequired;
                worksheet.Cell(row, 5).Value = stat.TotalScanned;
                worksheet.Cell(row, 6).Value = stat.Progress / 100;
                worksheet.Cell(row, 6).Style.NumberFormat.Format = "0.0%";

                if (row % 2 == 0)
                {
                    worksheet.Range($"A{row}:F{row}").Style.Fill.BackgroundColor = XLColor.FromHtml("#F8F9FA");
                }

                row++;
            }

            worksheet.Columns().AdjustToContents();
        }
    }
}