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

      // request 1 - get trending files around a specific user (me)
      // var request = client.Me.Insights.Trending.Request();

      // var results = request.GetAsync().Result;
      // foreach (var resource in results)
      // {
      //   Console.WriteLine("(" + resource.ResourceVisualization.Type + ") - " + resource.ResourceVisualization.Title);
      //   Console.WriteLine("  Weight: " + resource.Weight);
      //   Console.WriteLine("  Id: " + resource.Id);
      //   Console.WriteLine("  ResourceId: " + resource.ResourceReference.Id);
      // }

      // request 2 - used files
      var request = client.Me.Insights.Used.Request();

      var results = request.GetAsync().Result;
      foreach (var resource in results)
      {
        Console.WriteLine("(" + resource.ResourceVisualization.Type + ") - " + resource.ResourceVisualization.Title);
        Console.WriteLine("  Last Accessed: " + resource.LastUsed.LastAccessedDateTime.ToString());
        Console.WriteLine("  Last Modified: " + resource.LastUsed.LastModifiedDateTime.ToString());
        Console.WriteLine("  Id: " + resource.Id);
        Console.WriteLine("  ResourceId: " + resource.ResourceReference.Id);
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