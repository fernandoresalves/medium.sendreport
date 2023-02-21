using ClosedXML.Excel;
using Azure.Storage.Blobs.Specialized;
using Azure.Storage.Blobs;

//INFO: Configure Azure Storage connection strings
//REF: https://learn.microsoft.com/en-us/azure/storage/common/storage-configure-connection-string
var connectionString = "connection-string";
var blobContainerName = "name-container";

var container = new BlobContainerClient(connectionString, blobContainerName);
var client = container.GetBlockBlobClient($"relatorio-{DateTime.Now.ToString("yyyyMMdd")}.xlsx");

using (var workbook = new XLWorkbook())
{
    var worksheet = workbook.Worksheets.Add("Nome Aba");

    worksheet.Cell(1,1).Value = "Relátorio Trimestral";
    worksheet.Cell(1, 2).FormulaA1 = "=UPPER(MID(A1; 1; 9))";

    using MemoryStream memory = new();
    workbook.SaveAs(memory);
    memory.Position = 0;

    await client.UploadAsync(memory);
}
