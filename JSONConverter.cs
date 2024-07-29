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
using Azure.Identity;
using Azure.Storage.Blobs;


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
                var storageAccountUrl = req.Form["storageAccountUrl"];
                var containerName = req.Form["containerName"];
                var blobName = req.Form["blobName"];

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

                // Converts file to json
                DataSet ds = ExcelToDataSet(data: ms, hasHeader: true);
                var finalJson = DataTableToJSON(ds.Tables[0]);
                    
                // Create stream from the json
                using var fileStream = new MemoryStream(Encoding.UTF8.GetBytes(finalJson));

                // Create a BlobServiceClient using DefaultAzureCredential for managed identity authentication
                var blobServiceClient = new BlobServiceClient(new Uri(storageAccountUrl), new DefaultAzureCredential());
                // Get a reference to the Blob container
                var containerClient = blobServiceClient.GetBlobContainerClient(containerName);
                // Ensure the container exists
                await containerClient.CreateIfNotExistsAsync();
                // Get a reference to the Blob
                var blobClient = containerClient.GetBlobClient(blobName);
                // Upload the JSON file to the Blob
                await blobClient.UploadAsync(fileStream, true);
                
                return new OkObjectResult($"Finished successfully. JSON file pushed: {blobName}");
            }
            catch (Exception e)
            {
                return new OkObjectResult($"Blob Service Endpoint values are: {storageAccountUrl}, {containerName}, {blobName}" + e.Message);
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
                    var value = Convert.ToString(row[col]);
                    
                    if (!string.IsNullOrEmpty(value)) // Check if the column has a value
                    {
                        dict[col.ColumnName] = value;
                    }
                }
                list.Add(dict);
            }
            // Define serialization options with indentation
            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };
        
            // Serialize the dictionary to JSON using System.Text.Json with indentation
            return JsonSerializer.Serialize(list, options);
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
