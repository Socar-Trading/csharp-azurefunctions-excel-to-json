using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using ExcelDataReader;
using ChoETL;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Linq;

namespace br.feevale
{
    public static class JSONConverter
    {
        /// <summary>
        /// Function body.
        /// </summary>
        [FunctionName("JSONConverter")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            // Registers the Win1252 encoder used by Excel files.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            try
            {
                var formData = await req.ReadFormAsync();
                var file = req.Form.Files[0];

                // Checks the type of the file
                UploadedFileType filetype = UploadedFileType.INVALID;

                var extension = Path.GetExtension(file.FileName);
                if (extension == ".xlsx" || extension == ".xls")
                    filetype = UploadedFileType.EXCEL;

                if (extension == ".csv")
                    filetype = UploadedFileType.CSV;

                if(filetype == UploadedFileType.INVALID)
                    throw new FormatException("Invalid file type.");

                // Reads the file into memory.
                using var ms = new MemoryStream();
                await file.CopyToAsync(ms);
                ms.Seek(0, SeekOrigin.Begin);

                // Converts the stream to JSON
                using Stream stream = file.OpenReadStream();
                string result = string.Empty;

                if(filetype == UploadedFileType.EXCEL)
                {
                    DataSet ds = ExcelToDataSet(data: ms, hasHeader: true);
                    result = ds.Tables[0].ToCSV().ToJSON();
                }
                else
                {
                    DataSet ds = CSVToDataSet(data: ms, hasHeader: true);
                    result = ds.Tables[0].ToCSV().ToJSON();
                }

                // Returns the JSON content.
                return new OkObjectResult(result);
            }
            catch (Exception e)
            {
                return new OkObjectResult(e);
            }

        }

        /// <summary>
        /// Identifies the type of the uploaded file.
        /// </summary>
        private enum UploadedFileType
        {
            CSV,
            EXCEL,
            INVALID
        }

        /// <summary>
        /// Converts a DataTable into a CSV string.
        /// </summary>
        public static string ToCSV(this DataTable dtDataTable)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                // Headers
                IEnumerable<string> columnNames 
                    = dtDataTable.Columns.Cast<DataColumn>().
                        Select(column => $"\"{column.ColumnName.Trim()}\"");

                sb.AppendLine(string.Join(",", columnNames));

                // Lines
                foreach (DataRow row in dtDataTable.Rows)
                {
                    IEnumerable<string> fields = row.ItemArray.Select(field =>
                      string.Concat("\"", field.ToString().Trim().Replace("\"", "\"\""), "\""));

                    sb.AppendLine(string.Join(",", fields));
                }

                return sb.ToString();
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Converts a CSV string into JSON.
        /// </summary>
        public static string ToJSON(this string source)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                using var w = new ChoJSONWriter(sb);

                // Sets up the conversion schema.
                // Assumes that the first row is a header
                foreach (dynamic rec in ChoCSVReader
                    .LoadText(source)
                    .Configure(c => {
                        c.Delimiter = ",";
                        c.AutoDiscoverColumns = true;
                        c.AutoDiscoverFieldTypes = true;
                        c.ThrowAndStopOnMissingField = false;
                        c.MayContainEOLInData = true;
                        c.NullValueHandling = ChoNullValueHandling.Empty;
                        c.QuoteAllFields = true;
                })
                    .WithFirstLineHeader())
                {
                    w.Write(rec);
                }

                w.Close();

                return sb.ToString();
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Converts a CSV to a DataSet.
        /// </summary>
        public static DataSet CSVToDataSet(MemoryStream data, bool hasHeader = false)
        {
            try
            {
                using var reader = ExcelReaderFactory.CreateCsvReader(data);
                return reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                    {
                        EmptyColumnNamePrefix = "Column",
                        UseHeaderRow = hasHeader,
                    }
                });
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Converts an Excel file (.xls) to a DataSet.
        /// </summary>
        public static DataSet ExcelToDataSet(MemoryStream data, bool hasHeader = false)
        {
            try
            {
                using var reader = ExcelReaderFactory.CreateReader(data);
                return reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                    {
                        EmptyColumnNamePrefix = "Column",
                        UseHeaderRow = hasHeader,
                    },
                    UseColumnDataType = true
                });
            }
            catch
            {
                throw;
            }
        }

    }
}
