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

      // request 1 - get user's files
      // var request = client.Me.Drive.Root.Children.Request();

      // var results = request.GetAsync().Result;
      // foreach (var file in results)
      // {
      //   Console.WriteLine(file.Id + ": " + file.Name);
      // }

      // request 2 - get specific file
      // var fileId = "01X5UGSE5GWHXARY3YPRE3N4RAHPCJ7K3L";
      // var request = client.Me.Drive.Items[fileId].Request();

      // var results = request.GetAsync().Result;
      // Console.WriteLine(results.Id + ": " + results.Name);

      // request 3 - download specific file
      var fileId = "01X5UGSE5GWHXARY3YPRE3N4RAHPCJ7K3L";
      var request = client.Me.Drive.Items[fileId].Content.Request();

      var stream = request.GetAsync().Result;
      var driveItemPath = Path.Combine(System.IO.Directory.GetCurrentDirectory(), "driveItem_" + fileId + ".file");
      var driveItemFile = System.IO.File.Create(driveItemPath);
      stream.Seek(0, SeekOrigin.Begin);
      stream.CopyTo(driveItemFile);
      Console.WriteLine("Saved file to: " + driveItemPath);
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