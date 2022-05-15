// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;

namespace graphconsoleapp
{
  public class Program
  {
    public static void Main(string[] args)
    {
      var config = LoadAppSettings();
      if (config == null)
      {
        Console.WriteLine("Invalid appsettings.json file.");
        return;
      }

      var client = GetAuthenticatedGraphClient(config);

      var profileResponse = client.Me.Request().GetAsync().Result;
      Console.WriteLine("Hello " + profileResponse.DisplayName);

      // request 1 - upload small file to user's onedrive
      // var fileName = "smallfile.txt";
      // var filePath = Path.Combine(System.IO.Directory.GetCurrentDirectory(), fileName);
      // Console.WriteLine("Uploading file: " + fileName);

      // FileStream fileStream = new FileStream(filePath, FileMode.Open);
      // var uploadedFile = client.Me.Drive.Root
      //                               .ItemWithPath("smallfile.txt")
      //                               .Content
      //                               .Request()
      //                               .PutAsync<DriveItem>(fileStream)
      //                               .Result;
      // Console.WriteLine("File uploaded to: " + uploadedFile.WebUrl);

      // request 2 - upload large file to user's onedrive
      var fileName = "largefile.zip";
      var filePath = Path.Combine(System.IO.Directory.GetCurrentDirectory(), fileName);
      Console.WriteLine("Uploading file: " + fileName);

      // load resource as a stream
      using (Stream stream = new FileStream(filePath, FileMode.Open))
      {
        var uploadSession = client.Me.Drive.Root
                                        .ItemWithPath(fileName)
                                        .CreateUploadSession()
                                        .Request()
                                        .PostAsync()
                                        .Result;

        // create upload task
        var maxChunkSize = 320 * 1024;
        var largeUploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, stream, maxChunkSize);

        // create progress implementation
        IProgress<long> uploadProgress = new Progress<long>(uploadBytes =>
        {
          Console.WriteLine($"Uploaded {uploadBytes} bytes of {stream.Length} bytes");
        });

        // upload file
        UploadResult<DriveItem> uploadResult = largeUploadTask.UploadAsync(uploadProgress).Result;
        if (uploadResult.UploadSucceeded)
        {
          Console.WriteLine("File uploaded to user's OneDrive root folder.");
        }
      }
    }

    private static IConfigurationRoot? LoadAppSettings()
    {
      try
      {
        var config = new ConfigurationBuilder()
                          .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                          .AddJsonFile("appsettings.json", false, true)
                          .Build();

        if (string.IsNullOrEmpty(config["applicationId"]) ||
            string.IsNullOrEmpty(config["tenantId"]))
        {
          return null;
        }

        return config;
      }
      catch (System.IO.FileNotFoundException)
      {
        return null;
      }
    }

    private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
    {
      var clientId = config["applicationId"];
      var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

      List<string> scopes = new List<string>();
      scopes.Add("https://graph.microsoft.com/.default");

      var cca = PublicClientApplicationBuilder.Create(clientId)
                                              .WithAuthority(authority)
                                              .WithDefaultRedirectUri()
                                              .Build();
      return MsalAuthenticationProvider.GetInstance(cca, scopes.ToArray());
    }

    private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
    {
      var authenticationProvider = CreateAuthorizationProvider(config);
      var graphClient = new GraphServiceClient(authenticationProvider);
      return graphClient;
    }
  }
}