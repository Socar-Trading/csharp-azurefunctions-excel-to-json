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
using System.Web.Script.Serialization;



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

                // Converts the stream to JSON
                using Stream stream = file.OpenReadStream();
                string result = string.Empty;

                DataSet ds = ExcelToDataSet(data: ms, hasHeader: true);
                result = ds.Tables[0].DataTableToJSON();

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
        /// Converts DataTable into JSON.
        /// </summary>
        public static string DataTableToJSON(DataTable table)
        {
            // Initialize a dictionary to store column names and their corresponding values
            Dictionary<string, List<object>> dict = new Dictionary<string, List<object>>();
        
            // Loop through each column in the table
            foreach (DataColumn col in table.Columns)
            {
                // Initialize a list to store values for the current column
                List<object> columnValues = new List<object>();
        
                // Loop through each row in the table
                foreach (DataRow row in table.Rows)
                {
                    // Add the current cell value to the column's list
                    columnValues.Add(Convert.ToString(row[col]));
                }
        
                // Add the column and its values to the dictionary
                dict[col.ColumnName] = columnValues;
            }
        
            // Serialize the dictionary to JSON
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            return serializer.Serialize(dict);
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
