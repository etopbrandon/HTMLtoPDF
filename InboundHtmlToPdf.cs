using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.AspNetCore.Http;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using PuppeteerSharp;
using CreateUploadSessionPostRequestBody = Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession.CreateUploadSessionPostRequestBody;

namespace HTMLtoPDF;

public static class InboundHtmlToPdf
{
    [FunctionName("InboundHtmlToPdf")]
    public static async Task<HttpResponseMessage> RunAsync(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req, ILogger log)
    {
        log.LogInformation("Received Request for new Document");
        log.LogInformation($"Environment {Environment.GetEnvironmentVariable("AZURE_FUNCTIONS_ENVIRONMENT")}");
        string incomingBody = await new StreamReader(req.Body).ReadToEndAsync();
        HtmlPost html = JsonConvert.DeserializeObject<HtmlPost>(incomingBody);

        var browserlessApiKey = Environment.GetEnvironmentVariable("browserlessApiKey");

        var options = new ConnectOptions()
        {
            BrowserWSEndpoint = $"wss://chrome.browserless.io?token={browserlessApiKey}"
        };
        var browser = await Puppeteer.ConnectAsync(options);

        var page = await browser.NewPageAsync();

        await page.SetContentAsync(html?.Html);
        log.LogInformation("Converting PDF");
        var pdfStream = await page.PdfStreamAsync();
        log.LogInformation("PDF Stream Captured");
        await browser.CloseAsync();
        log.LogInformation("Getting Graph Client");
        var client = GetGraph();
        log.LogInformation("Starting File Tasks (Upload and Encoding)");
        log.LogInformation("Creating Byte Array");
        byte[] pdfArray;
        log.LogInformation("Adding to array from Memory Stream");
        using (var ms = new MemoryStream())
        {
            await pdfStream.CopyToAsync(ms);
            pdfArray = ms.ToArray();
        }
        log.LogInformation("Starting File Upload");
        var uploadResponse = await UploadFile(client, pdfStream, html?.ClientName, log);
        log.LogInformation("File Task Completed");
        var apiResponse = new ApiResponse();
        if (pdfArray.Length > 0 && uploadResponse.success)
        {
            apiResponse.base64 = Convert.ToBase64String(pdfArray);
            apiResponse.uploadUrl = uploadResponse.response;
            apiResponse.success = true;
            log.LogInformation("Conversion Succeeded");
        }
        else
        {
            apiResponse.success = false;
            apiResponse.uploadErrors = uploadResponse.response;
            log.LogError("Conversion FAILED! See API Response");
        }

        var jsonResponse = JsonConvert.SerializeObject(apiResponse);
        
        log.LogInformation("Sending Response");
        return new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(jsonResponse, Encoding.UTF8, "application/json")
        };
    }

    private static GraphServiceClient GetGraph()
    {
        ChainedTokenCredential credential;
        if (Environment.GetEnvironmentVariable("AZURE_FUNCTIONS_ENVIRONMENT") == "Development")
        {
            credential = new ChainedTokenCredential(new ClientSecretCredential(
                tenantId: Environment.GetEnvironmentVariable("tenantId"),
                clientId: Environment.GetEnvironmentVariable("clientId"),
                clientSecret: Environment.GetEnvironmentVariable("clientSecret")));
            Console.WriteLine("Development Creds Chosen");
        }
        else
        {
            credential = new ChainedTokenCredential(new ManagedIdentityCredential());
            Console.WriteLine("Production Creds Chosen");
        }

        //var credential = new EnvironmentCredential();
        string[] scopes = { "https://graph.microsoft.com/.default" };

        var client = new GraphServiceClient(credential, scopes);
        return client;
    }

    private static async Task<UploadResponse> UploadFile(GraphServiceClient client, Stream pdfStream, string clientName, ILogger log)
    {
        //TODO: /Teams[group-id]/channels[channel-id]/filesFolder
        //Returns driveId and parentId
        //pull those out too?
        var driveId = Environment.GetEnvironmentVariable("driveId");
        var parentId = Environment.GetEnvironmentVariable("parentId");
        var fileName = $"{clientName}{DateTime.Now:yyyyMMdd}BECReport.pdf";
        var uploadResponseReturn = new UploadResponse();
        var uploadProps = new CreateUploadSessionPostRequestBody()
        {
            AdditionalData = new Dictionary<string, object>
            {
                { "@microsoft.graph.conflictBehavior", "rename" }
            }
        };
        // POST /drives/{driveId}/items/{itemId}/createUploadSession
        UploadSession uploadSession = null;
        try
        {
            uploadSession = await client.Drives[driveId].Items[parentId]
                .ItemWithPath(fileName)
                .CreateUploadSession
                .PostAsync(uploadProps);
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            uploadResponseReturn.response = e.ToString();
            uploadResponseReturn.success = false;
        }
        var maxSliceSize = 320 * 1024;
        var fileUploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, pdfStream, maxSliceSize);
        var uploadResponse = new UploadResult<DriveItem>();
        try
        {
            uploadResponse = await fileUploadTask.UploadAsync();
        }
        catch (Exception e)
        {
            uploadResponseReturn.response = e.ToString();
            uploadResponseReturn.success = false;
        }

        if (uploadResponse.UploadSucceeded)
        {
            uploadResponseReturn.response = uploadResponse.ItemResponse.WebUrl;
            uploadResponseReturn.success = true;
        }

        return uploadResponseReturn;
    }
}

public class HtmlPost
{
    public string Html { get; set; }
    public string ClientName { get; set; }
}

public class ApiResponse
{
    public string base64 { get; set; }
    public bool success { get; set; }
    public string uploadUrl { get; set; }
    public string uploadErrors { get; set; }
}

public class UploadResponse
{
    public bool success { get; set; }
    public string response { get; set; }
}