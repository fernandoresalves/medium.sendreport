//https://github.com/ClosedXML/ClosedXML
using ClosedXML.Excel;

//https://github.com/Azure/azure-sdk-for-net/tree/main/sdk/storage/Azure.Storage.Blobs
using Azure.Storage.Blobs.Specialized;
using Azure.Storage.Blobs;

var connectionString = "connection-string";
var blobContainerName = "name-container";

var container = new BlobContainerClient(connectionString, blobContainerName);
var client = container.GetBlockBlobClient($"relatorio-{DateTime.Now.ToString("yyyyMMdd")}.xlsx");

using (var workbook = new XLWorkbook())
{
    var worksheet = workbook.Worksheets.Add("Nome Aba");

    worksheet.Cell(1,1).Value = "Relátorio Trimestral";

    //https://support.google.com/docs/answer/3094219?hl=en&ref_topic=3105625
    //https://support.google.com/docs/answer/3094129?hl=en&ref_topic=3105625

    worksheet.Cell(1, 2).FormulaA1 = "=UPPER(MID(A1; 1; 9))";

    //https://learn.microsoft.com/pt-br/dotnet/api/system.io.memorystream?view=net-7.0
    using MemoryStream memory = new();
    workbook.SaveAs(memory);
    memory.Position = 0;

    await client.UploadAsync(memory);

}
