// <copyright file="ICompanyResponseSearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// Interface which helps to search for company response using Azure Search service.
    /// </summary>
    public interface ICompanyResponseSearchService
    {
        /// <summary>
        /// Provide search result for table to be used by SME based on Azure Search service.
        /// </summary>
        /// <param name="searchQuery">Query which the user had typed in messaging extension search field.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <param name="isTaskModuleData">Boolean value which indicates to get the data for company responses task module or not.</param>
        /// <returns>List of search company response results.</returns>
        Task<IList<CompanyResponseEntity>> GetSearchCompanyResponseAsync(string searchQuery, int? count = null, int? skip = null, bool isTaskModuleData = false);
    }
}
