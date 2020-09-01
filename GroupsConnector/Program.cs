// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using GroupsConnector.Authentication;
using GroupsConnector.Console;
using GroupsConnector.Graph;
using GroupsConnector.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace GroupsConnector
{
    class Program
    {
        private static GraphHelper _graphHelper;

        private static ExternalConnection _currentConnection;

        private static IConfigurationRoot _appConfig;

        static async Task Main(string[] args)
        {
            try
            {
                Output.WriteLine("Groups Search Connector\n");

                // Load configuration from appsettings.json
                _appConfig = LoadAppSettings();
                if (_appConfig == null)
                {
                    Output.WriteLine(Output.Error, "Missing or invalid user secrets");
                    Output.WriteLine(Output.Error, "Please see README.md for instructions on configuring the application.");
                    return;
                }

                // Initialize the auth provider
                var authProvider = new ClientCredentialAuthProvider(
                    _appConfig["appId"],
                    _appConfig["tenantId"],
                    _appConfig["appSecret"]
                );

                _graphHelper = new GraphHelper(authProvider);

                do
                {
                    var userChoice = DoMenuPrompt();

                    switch (userChoice)
                    {
                        case MenuChoice.CreateConnection:
                            await CreateConnectionAsync();
                            break;
                        case MenuChoice.ChooseExistingConnection:
                            await SelectExistingConnectionAsync();
                            break;
                        case MenuChoice.DeleteConnection:
                            await DeleteCurrentConnectionAsync();
                            break;
                        case MenuChoice.RegisterSchema:
                            await RegisterSchemaAsync();
                            break;
                        case MenuChoice.ViewSchema:
                            await GetSchemaAsync();
                            break;
                        case MenuChoice.PushUpdatedItems:
                            await UpdateItems(true);
                            break;
                        case MenuChoice.PushAllItems:
                            await UpdateItems(false);
                            break;
                        case MenuChoice.Exit:
                            // Exit the program
                            Output.WriteLine("Goodbye...");
                            return;
                        case MenuChoice.Invalid:
                        default:
                            Output.WriteLine(Output.Warning, "Invalid choice! Please try again.");
                            break;
                    }

                    Output.WriteLine("");

                } while (true);
            }
            catch (Exception ex)
            {
                Output.WriteLine(Output.Error, "An unexpected exception occurred.");
                Output.WriteLine(Output.Error, ex.Message);
                Output.WriteLine(Output.Error, ex.StackTrace);
            }
        }

        private static async Task CreateConnectionAsync()
        {
            var connectionId = PromptForInput("Enter a unique ID for the new connection", true);
            var connectionName = PromptForInput("Enter a name for the new connection", true);
            var connectionDescription = PromptForInput("Enter a description for the new connection", false);

            try
            {
                // Create the connection
                _currentConnection = await _graphHelper.CreateConnectionAsync(connectionId, connectionName, connectionDescription);
                Output.WriteLine(Output.Success, "New connection created");
                Output.WriteObject(Output.Info, _currentConnection);
            }
            catch (ServiceException serviceException)
            {
                Output.WriteLine(Output.Error, $"{serviceException.StatusCode} error creating new connection:");
                Output.WriteLine(Output.Error, serviceException.Message);
                return;
            }
        }

        private static async Task SelectExistingConnectionAsync()
        {
            Output.WriteLine(Output.Info, "Getting existing connections...");
            try
            {
                // Get connections
                var connections = await _graphHelper.GetExistingConnectionsAsync();

                if (connections.CurrentPage.Count <= 0)
                {
                    Output.WriteLine(Output.Warning, "No connections exist. Please create a new connection.");
                    return;
                }

                Output.WriteLine(Output.Info, "Choose one of the following connections:");
                int menuNumber = 1;
                foreach(var connection in connections.CurrentPage)
                {
                    Output.WriteLine($"{menuNumber++}. {connection.Name}");
                }

                ExternalConnection selectedConnection = null;

                do
                {
                    try
                    {
                        Output.Write(Output.Info, "Selection: ");
                        var choice = int.Parse(System.Console.ReadLine());

                        if (choice > 0 && choice <= connections.CurrentPage.Count)
                        {
                            selectedConnection = connections.CurrentPage[choice-1];
                        }
                        else
                        {
                            Output.WriteLine(Output.Warning, "Invalid choice.");
                        }
                    }
                    catch (FormatException)
                    {
                        Output.WriteLine(Output.Warning, "Invalid choice.");
                    }
                } while (selectedConnection == null);

                _currentConnection = selectedConnection;
            }
            catch (ServiceException serviceException)
            {
                Output.WriteLine(Output.Error, $"{serviceException.StatusCode} error getting connections:");
                Output.WriteLine(Output.Error, serviceException.Message);
                return;
            }
        }

        private static async Task DeleteCurrentConnectionAsync()
        {
            if (_currentConnection == null)
            {
                Output.WriteLine(Output.Warning, "No connection selected. Please create a new connection or select an existing connection.");
                return;
            }

            Output.WriteLine(Output.Warning, $"Deleting {_currentConnection.Name} - THIS CANNOT BE UNDONE");
            Output.WriteLine(Output.Warning, "Enter the connection name to confirm.");

            var input = System.Console.ReadLine();

            if (input != _currentConnection.Name)
            {
                Output.WriteLine(Output.Warning, "Canceled");
            }

            try
            {
                await _graphHelper.DeleteConnectionAsync(_currentConnection.Id);
                Output.WriteLine(Output.Success, $"{_currentConnection.Name} deleted");
                _currentConnection = null;
            }
            catch (ServiceException serviceException)
            {
                Output.WriteLine(Output.Error, $"{serviceException.StatusCode} error deleting connection:");
                Output.WriteLine(Output.Error, serviceException.Message);
                return;
            }
        }

        private static async Task RegisterSchemaAsync()
        {
            if (_currentConnection == null)
            {
                Output.WriteLine(Output.Warning, "No connection selected. Please create a new connection or select an existing connection.");
                return;
            }

            Output.WriteLine(Output.Info, "Registering schema, this may take a moment...");

            try
            {
                // Register the schema
                var schema = new Schema
                {
                    // Need to set to null, service returns 400
                    // if @odata.type property is sent
                    ODataType = null,
                    BaseType = "microsoft.graph.externalItem",
                    Properties = new List<Property>
                    {
                        new Property { Name = "id", Type = PropertyType.String, IsQueryable = true, IsSearchable = false, IsRetrievable = true },
                        new Property { Name = "displayName", Type = PropertyType.String, IsQueryable = true, IsSearchable = true, IsRetrievable = true },
                        new Property { Name = "description", Type = PropertyType.String, IsQueryable = false, IsSearchable = true, IsRetrievable = true },
                    }
                };

                await _graphHelper.RegisterSchemaAsync(_currentConnection.Id, schema);
                Output.WriteLine(Output.Success, "Schema registered");
            }
            catch (ServiceException serviceException)
            {
                Output.WriteLine(Output.Error, $"{serviceException.StatusCode} error registering schema:");
                Output.WriteLine(Output.Error, serviceException.Message);
                return;
            }
        }

        private static async Task GetSchemaAsync()
        {
            if (_currentConnection == null)
            {
                Output.WriteLine(Output.Warning, "No connection selected. Please create a new connection or select an existing connection.");
                return;
            }

            try
            {
                var schema = await _graphHelper.GetSchemaAsync(_currentConnection.Id);
                Output.WriteObject(Output.Info, schema);
            }
            catch (ServiceException serviceException)
            {
                Output.WriteLine(Output.Error, $"{serviceException.StatusCode} error getting schema:");
                Output.WriteLine(Output.Error, serviceException.Message);
                return;
            }
        }

        private static async Task UpdateItems(bool uploadModifiedOnly)
        {
            if (_currentConnection == null)
            {
                Output.WriteLine(Output.Warning, "No connection selected. Please create a new connection or select an existing connection.");
                return;
            }

            List<AADGroup> groupsToUpload = null;
            List<AADGroup> groupsToDelete = null;

            var newUploadTime = DateTime.UtcNow;

            // Get all AAD Groups
            var groups = await _graphHelper.GetAllGroups();

            if (uploadModifiedOnly)
            {
                // Load the last upload timestamp
                var lastUploadTime = GetLastUploadTime();
                Output.WriteLine(Output.Info, $"Uploading changes since last upload at {lastUploadTime.ToLocalTime().ToString()}");

                groupsToUpload = groups
                    .Where(p => p.DeletedDateTime == DateTime.MinValue && p.CreatedDateTime > lastUploadTime)
                    .ToList();

                groupsToDelete = groups
                    .Where(p => p.DeletedDateTime != DateTime.MinValue && p.DeletedDateTime > lastUploadTime)
                    .ToList();            
            }
            else
            {
                groupsToUpload = groups
                    .Where(p => p.DeletedDateTime == DateTime.MinValue)
                    .ToList();

                groupsToDelete = new List<AADGroup>();
            }            

            Output.WriteLine(Output.Info, $"Processing {groupsToUpload.Count()} add/updates, {groupsToDelete.Count()} deletes");
            bool success = true;

            foreach(var group in groupsToUpload)
            {
                var newItem = new ExternalItem
                {
                    Id = group.Id.ToString(),
                    Content = new ExternalItemContent
                    {
                        // Need to set to null, service returns 400
                        // if @odata.type property is sent
                        ODataType = null,
                        Type = ExternalItemContentType.Text,
                        Value = group.Description
                    },
                    Acl = new List<Acl>
                    {
                        new Acl {
                            AccessType = AccessType.Grant,
                            Type = AclType.Group,
                            Value = group.Id,
                            IdentitySource = "Azure Active Directory"
                        }
                    },
                    Properties = group.AsExternalItemProperties()
                };

                try
                {
                    Output.Write(Output.Info, $"Uploading group {group.DisplayName} ({group.Id})...");
                    await _graphHelper.AddOrUpdateItem(_currentConnection.Id, newItem);
                    Output.WriteLine(Output.Success, "DONE");
                }
                catch (ServiceException serviceException)
                {
                    success = false;
                    Output.WriteLine(Output.Error, "FAILED");
                    Output.WriteLine(Output.Error, $"{serviceException.StatusCode} error adding or updating group {group.Id}");
                    Output.WriteLine(Output.Error, serviceException.Message);
                }
            }

            foreach (var part in groupsToDelete)
            {
                try
                {
                    Output.Write(Output.Info, $"Deleting part number {part.Id}...");
                    await _graphHelper.DeleteItem(_currentConnection.Id, part.Id.ToString());
                    Output.WriteLine(Output.Success, "DONE");
                }
                catch (ServiceException serviceException)
                {
                    if (serviceException.StatusCode.Equals(System.Net.HttpStatusCode.NotFound))
                    {
                        Output.WriteLine(Output.Warning, "Not found");
                    }
                    else
                    {
                        success = false;
                        Output.WriteLine(Output.Error, "FAILED");
                        Output.WriteLine(Output.Error, $"{serviceException.StatusCode} error deleting part {part.Id}");
                        Output.WriteLine(Output.Error, serviceException.Message);
                    }
                }
            }

            // If no errors, update our last upload time
            if (success)
            {
                SaveLastUploadTime(newUploadTime);
            }
        }

        private static readonly string uploadTimeFile = "lastuploadtime.bin";

        private static DateTime GetLastUploadTime()
        {
            if (System.IO.File.Exists(uploadTimeFile))
            {
                var uploadTimeString = System.IO.File.ReadAllText(uploadTimeFile);
                return DateTime.Parse(uploadTimeString).ToUniversalTime();
            }

            return DateTime.MinValue;
        }

        private static void SaveLastUploadTime(DateTime uploadTime)
        {
            System.IO.File.WriteAllText(uploadTimeFile, uploadTime.ToString("u"));
        }

        private static MenuChoice DoMenuPrompt()
        {
            Output.WriteLine(Output.Info, $"Current connection: {(_currentConnection == null ? "NONE" : _currentConnection.Name)}");
            Output.WriteLine(Output.Info, "Please choose one of the following options:");

            Output.WriteLine($"{Convert.ToInt32(MenuChoice.CreateConnection)}. Create a connection");
            Output.WriteLine($"{Convert.ToInt32(MenuChoice.ChooseExistingConnection)}. Select an existing connection");
            Output.WriteLine($"{Convert.ToInt32(MenuChoice.DeleteConnection)}. Delete current connection");
            Output.WriteLine($"{Convert.ToInt32(MenuChoice.RegisterSchema)}. Register schema for current connection");
            Output.WriteLine($"{Convert.ToInt32(MenuChoice.ViewSchema)}. View schema for current connection");
            Output.WriteLine($"{Convert.ToInt32(MenuChoice.PushUpdatedItems)}. Push updated items to current connection");
            Output.WriteLine($"{Convert.ToInt32(MenuChoice.PushAllItems)}. Push ALL items to current connection");
            Output.WriteLine($"{Convert.ToInt32(MenuChoice.Exit)}. Exit");

            try
            {
                var choice = int.Parse(System.Console.ReadLine());
                return (MenuChoice)choice;
            }
            catch (FormatException)
            {
                return MenuChoice.Invalid;
            }
        }

        private static string PromptForInput(string prompt, bool valueRequired)
        {
            string response = null;

            do
            {
                Output.WriteLine(Output.Info, $"{prompt}:");
                response = System.Console.ReadLine();
                if (valueRequired && string.IsNullOrEmpty(response))
                {
                    Output.WriteLine(Output.Error, "You must provide a value");
                }
            } while (valueRequired && string.IsNullOrEmpty(response));

            return response;
        }

        private static IConfigurationRoot LoadAppSettings()
        {
            var appConfig = new ConfigurationBuilder()
                .AddUserSecrets<Program>()
                .Build();

            // Check for required settings
            if (string.IsNullOrEmpty(appConfig["appId"]) ||
                string.IsNullOrEmpty(appConfig["appSecret"]) ||
                string.IsNullOrEmpty(appConfig["tenantId"]))
            {
                return null;
            }

            return appConfig;
        }
    }
}
