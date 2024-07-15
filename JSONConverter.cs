using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using ExcelDataReader;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;



namespace JSONConverter
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
                
                DataSet ds = ExcelToDataSet(data: ms, hasHeader: true);
                // Returns the JSON content.
                return new OkObjectResult(DataTableToJSON(ds.Tables[0]));
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
        /// Converts DataTable into JSON.
        /// </summary>
        public static string DataTableToJSON(DataTable table)
        {
            List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();

            foreach (DataRow row in table.Rows)
            {
                Dictionary<string, object> dict = new Dictionary<string, object>();

                foreach (DataColumn col in table.Columns)
                {
                    dict[col.ColumnName] = (Convert.ToString(row[col]));
                }
                list.Add(dict);
            }
            // Define serialization options with indentation
            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };
        
            // Serialize the dictionary to JSON using System.Text.Json with indentation
            return JsonSerializer.Serialize(dict, options);
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
