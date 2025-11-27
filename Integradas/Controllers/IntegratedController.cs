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

                order.ScannedQuantity = (order.ScannedQuantity ?? 0) + 1;

                var amount = order.Amount ?? 0;
                if (order.ScannedQuantity >= amount)
                {
                    order.Completed = true;
                    order.CompletedDate = DateTime.UtcNow;
                }

                await _context.SaveChangesAsync();

                var remaining = Math.Max(0, amount - (int)order.ScannedQuantity);
                var progressPercentage = amount > 0 ?
                    (double)order.ScannedQuantity / amount * 100 : 0;

                var orderDto = new OrderDetailDto
                {
                    Id = order.Id,
                    Type = order.Type ?? string.Empty,
                    PartNumber = order.PartNumber ?? string.Empty,
                    Pipeline = order.Pipeline ?? string.Empty,
                    Amount = amount,
                    ScannedQuantity = order.ScannedQuantity,
                    Completed = order.Completed,
                    CreatedAt = order.CreatedAt,
                    CompletedDate = order.CompletedDate
                };

                return Ok(new ScanResponseDto
                {
                    Success = true,
                    Message = order.Completed ?
                        $"¡COMPLETADO! {order.PartNumber} - {order.ScannedQuantity}/{amount}" :
                        $"Escaneado: {order.ScannedQuantity}/{amount}",
                    PartNumber = order.PartNumber ?? string.Empty,
                    ScannedQuantity = order.ScannedQuantity,
                    RequiredQuantity = amount,
                    Completed = order.Completed,
                    Remaining = remaining,
                    ProgressPercentage = progressPercentage,
                    Order = orderDto
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en Scan: {ex.Message}");
                Console.WriteLine($"StackTrace: {ex.StackTrace}");

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
    }
}