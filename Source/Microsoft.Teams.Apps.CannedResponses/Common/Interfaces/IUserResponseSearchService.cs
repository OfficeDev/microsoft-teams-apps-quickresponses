// <copyright file="IUserResponseSearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// Interface for user response search service which helps in searching use response using Azure Search service.
    /// </summary>
    public interface IUserResponseSearchService
    {
        /// <summary>
        /// Provide search result for table to be used by SME based on Azure Search service.
        /// </summary>
        /// <param name="searchQuery">Query which the user had typed in messaging extension search field.</param>
        /// /// <param name="userObjectId">AAd object id of the user.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <returns>List of search results.</returns>
        Task<IList<UserResponseEntity>> SearchUserResponseAsync(string searchQuery, string userObjectId, int? count = null, int? skip = null);

        /// <summary>
        /// Creates Index, Data Source and Indexer for search service.
        /// </summary>
        /// <param name="connectionString">Connection string to the data store.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task InitializeSearchServiceIndexAsync(string connectionString);
    }
}
