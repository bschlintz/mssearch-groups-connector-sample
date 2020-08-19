// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.
using Microsoft.Graph;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace GroupsConnector.Models
{
    public class AADGroup
    {
        [Key]
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }

        public Properties AsExternalItemProperties()
        {
            var properties = new Properties
            {
                AdditionalData = new Dictionary<string, object>
                {
                    { "id", Id },
                    { "displayName", DisplayName },
                    { "description", Description },
                }
            };

            return properties;
        }
    }
}