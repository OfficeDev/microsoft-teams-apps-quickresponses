// <copyright file="SearchServiceSetting.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Models
{
    /// <summary>
    /// Provides settings related to search service.
    /// </summary>
    public class SearchServiceSetting : StorageSetting
    {
        /// <summary>
        /// Gets or sets search service name.
        /// </summary>
        public string SearchServiceName { get; set; }

        /// <summary>
        /// Gets or sets search service query api key.
        /// </summary>
        public string SearchServiceQueryApiKey { get; set; }

        /// <summary>
        /// Gets or sets search service admin api key.
        /// </summary>
        public string SearchServiceAdminApiKey { get; set; }

        /// <summary>
        /// Gets or sets search indexing interval in minutes.
        /// </summary>
        public string SearchIndexingIntervalInMinutes { get; set; }
    }
}
