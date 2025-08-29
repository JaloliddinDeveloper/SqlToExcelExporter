using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using SqlToExcelExporter.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace SqlToExcelExporter.Controllers
{
    // 🔸 Yangi model — Kategoriyali query
    public class QueryItem
    {
        public string Name { get; set; }
        public string Sql { get; set; }
        public string Category { get; set; }
    }

    // 🔸 O‘chirish uchun model
    public class DeleteQueryRequest
    {
        public string Name { get; set; }
        public string Category { get; set; }
    }

    public class ExportController : Controller
    {
        private static int ProgressCount = 0;
        private static bool ExportDone = false;
        private static int TotalRows = 0;

        private readonly IWebHostEnvironment _env;
        private string QueriesFile => Path.Combine(_env.WebRootPath, "queries.json");

        public ExportController(IWebHostEnvironment env)
        {
            _env = env;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View(new ExportRequest());
        }

        [HttpGet]
        public IActionResult GetProgress()
        {
            return Json(new { rows = ProgressCount, done = ExportDone, total = TotalRows });
        }

        [HttpGet]
        public IActionResult GetSavedQueries()
        {
            if (!System.IO.File.Exists(QueriesFile))
                return Json(new List<QueryItem>());

            var json = System.IO.File.ReadAllText(QueriesFile);
            var queries = JsonSerializer.Deserialize<List<QueryItem>>(json);
            return Json(queries ?? new List<QueryItem>());
        }

        [HttpPost]
        public IActionResult SaveQuery([FromBody] QueryItem query)
        {
            if (string.IsNullOrWhiteSpace(query?.Name) || string.IsNullOrWhiteSpace(query.Sql))
                return BadRequest("Query nomi va SQL bo‘sh bo‘lmasligi kerak");

            List<QueryItem> queries = new List<QueryItem>();
            if (System.IO.File.Exists(QueriesFile))
            {
                var json = System.IO.File.ReadAllText(QueriesFile);
                queries = JsonSerializer.Deserialize<List<QueryItem>>(json) ?? new List<QueryItem>();
            }

            // Avvalgi query’ni yangilaymiz (name + category bo‘yicha)
            queries.RemoveAll(q =>
                string.Equals(q.Name, query.Name, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(q.Category ?? "", query.Category ?? "", StringComparison.OrdinalIgnoreCase));

            if (queries.Count >= 500) queries.RemoveAt(0);

            queries.Add(query);
            System.IO.File.WriteAllText(QueriesFile, JsonSerializer.Serialize(queries));

            return Ok(new { success = true });
        }

        [HttpPost]
        public IActionResult DeleteQuery([FromBody] DeleteQueryRequest req)
        {
            if (!System.IO.File.Exists(QueriesFile))
                return NotFound();

            var json = System.IO.File.ReadAllText(QueriesFile);
            var queries = JsonSerializer.Deserialize<List<QueryItem>>(json) ?? new List<QueryItem>();

            int removed = queries.RemoveAll(q =>
                string.Equals(q.Name, req.Name, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(q.Category ?? "", req.Category ?? "", StringComparison.OrdinalIgnoreCase));

            if (removed == 0)
                return NotFound();

            System.IO.File.WriteAllText(QueriesFile, JsonSerializer.Serialize(queries));
            return Ok(new { success = true });
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
            TotalRows = 0;
            List<byte[]> excelFiles = new List<byte[]>();

            using (SqlConnection conn = new SqlConnection(request.ConnectionString))
            {
                await conn.OpenAsync();

                using (SqlCommand countCmd = new SqlCommand($"SELECT COUNT(*) FROM ({request.SqlQuery}) AS t", conn))
                {
                    TotalRows = (int)await countCmd.ExecuteScalarAsync();
                }

                using (SqlCommand cmd = new SqlCommand(request.SqlQuery, conn))
                {
                    cmd.CommandTimeout = 600;
                    using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
                    {
                        int maxRowsPerFile = 1_000_000;
                        int rowCount = 0;
                        int fileIndex = 1;

                        MemoryStream ms = null;
                        SpreadsheetDocument document = null;
                        WorkbookPart workbookPart = null;
                        OpenXmlWriter writer = null;

                        List<string> headers = Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToList();

                        Action CloseCurrentFile = () =>
                        {
                            if (writer != null)
                            {
                                writer.WriteEndElement(); // SheetData
                                writer.WriteEndElement(); // Worksheet
                                writer.Close();
                                writer = null;
                            }

                            if (workbookPart != null)
                                workbookPart.Workbook.Save();

                            if (document != null)
                            {
                                document.Dispose();
                                document = null;
                            }

                            if (ms != null)
                            {
                                excelFiles.Add(ms.ToArray());
                                ms.Dispose();
                                ms = null;
                            }

                            fileIndex++;
                        };

                        void StartNewFile()
                        {
                            ms = new MemoryStream();
                            document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
                            workbookPart = document.AddWorkbookPart();
                            workbookPart.Workbook = new Workbook();
                            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                            writer = OpenXmlWriter.Create(worksheetPart);

                            writer.WriteStartElement(new Worksheet());
                            writer.WriteStartElement(new SheetData());

                            // Header yozamiz
                            WriteRow(writer, headers);

                            sheets.Append(new Sheet()
                            {
                                Id = workbookPart.GetIdOfPart(worksheetPart),
                                SheetId = 1,
                                Name = "Data"
                            });

                            rowCount = 0;
                        }

                        StartNewFile();

                        while (await reader.ReadAsync())
                        {
                            var values = new string[reader.FieldCount];
                            for (int i = 0; i < reader.FieldCount; i++)
                                values[i] = reader.IsDBNull(i) ? "" : reader[i].ToString();

                            WriteRow(writer, values);
                            rowCount++;
                            ProgressCount++;

                            if (rowCount >= maxRowsPerFile)
                            {
                                CloseCurrentFile();
                                StartNewFile();
                            }
                        }

                        CloseCurrentFile();
                        ExportDone = true;
                    }
                }
            }

            using (var zipStream = new MemoryStream())
            {
                using (var archive = new ZipArchive(zipStream, ZipArchiveMode.Create, true))
                {
                    for (int i = 0; i < excelFiles.Count; i++)
                    {
                        var entry = archive.CreateEntry($"{request.FileName}_{i + 1}.xlsx");
                        using (var entryStream = entry.Open())
                        {
                            entryStream.Write(excelFiles[i], 0, excelFiles[i].Length);
                        }
                    }
                }

                zipStream.Position = 0;
                return File(zipStream.ToArray(), "application/zip", request.FileName + ".zip");
            }
        }

        private void WriteRow(OpenXmlWriter writer, IEnumerable<string> values)
        {
            writer.WriteStartElement(new Row());
            foreach (var value in values)
            {
                writer.WriteElement(new Cell()
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(value ?? string.Empty)
                });
            }
            writer.WriteEndElement();
        }
    }
}
