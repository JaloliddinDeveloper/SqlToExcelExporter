using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using SqlToExcelExporter.Models;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace SqlToExcelExporter.Controllers
{
    public class ExportController : Controller
    {
        private static int ProgressCount = 0;
        private static bool ExportDone = false;

        [HttpGet]
        public IActionResult Index()
        {
            return View(new ExportRequest());
        }

        [HttpGet]
        public IActionResult GetProgress()
        {
            return Json(new { rows = ProgressCount, done = ExportDone });
        }

        [HttpPost]
        public async Task<IActionResult> Index(ExportRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.ConnectionString) ||
                string.IsNullOrWhiteSpace(request.SqlQuery) ||
                string.IsNullOrWhiteSpace(request.FileName))
            {
                ModelState.AddModelError("", "Barcha maydonlarni to‘ldiring!");
                return View(request);
            }

            ProgressCount = 0;
            ExportDone = false;

            using (SqlConnection conn = new SqlConnection(request.ConnectionString))
            {
                await conn.OpenAsync();
                using (SqlCommand cmd = new SqlCommand(request.SqlQuery, conn))
                using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
                using (var workbook = new XLWorkbook())
                {
                    int maxRowsPerSheet = 1_000_000;
                    int sheetIndex = 1;

                    var columnNames = Enumerable.Range(0, reader.FieldCount)
                                                .Select(reader.GetName)
                                                .ToList();

                    var rowsBuffer = new List<object[]>();
                    int totalRow = 0;

                    while (await reader.ReadAsync())
                    {
                        var row = new object[reader.FieldCount];
                        reader.GetValues(row);
                        rowsBuffer.Add(row);
                        totalRow++;

                        ProgressCount = totalRow; 

                        if (rowsBuffer.Count == maxRowsPerSheet)
                        {
                            WriteSheet(workbook, "Sheet" + sheetIndex, columnNames, rowsBuffer);
                            rowsBuffer.Clear();
                            sheetIndex++;
                        }
                    }

                    if (rowsBuffer.Count > 0)
                    {
                        WriteSheet(workbook, "Sheet" + sheetIndex, columnNames, rowsBuffer);
                    }

                    ExportDone = true;

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();
                        return File(content,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            request.FileName + ".xlsx");
                    }
                }
            }
        }

        private void WriteSheet(XLWorkbook workbook, string sheetName, List<string> headers, List<object[]> rows)
        {
            var ws = workbook.Worksheets.Add(sheetName);

            for (int i = 0; i < headers.Count; i++)
                ws.Cell(1, i + 1).Value = headers[i];

            ws.Cell(2, 1).InsertData(rows);
            ws.Columns().AdjustToContents();
        }
    }
}
