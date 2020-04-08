// <copyright file="ICompanyResponseStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// Interface for provider class which helps in storing company response details in Microsoft Azure Table storage.
    /// </summary>
    public interface ICompanyResponseStorageProvider
    {
        /// <summary>
        /// Get requests submitted by current user from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="userId">User Id of user for which requests needs to be fetched.</param>
        /// <returns>A task that represents a collection to hold user responses data.</returns>
        Task<IEnumerable<CompanyResponseEntity>> GetUserCompanyResponseAsync(string userId);

        /// <summary>
        /// Get company responses data from Microsoft Azure Table storage.
        /// </summary>
        /// <returns>A task that holds company response entity data in collection.</returns>
        Task<IEnumerable<CompanyResponseEntity>> GetCompanyResponsesDataAsync();

        /// <summary>
        /// Delete company response details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="entity">Holds company response detail entity data.</param>
        /// <returns>A task that represents company response entity data is saved or updated.</returns>
        Task<bool> DeleteEntityAsync(CompanyResponseEntity entity);

        /// <summary>
        /// Stores or update company response details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="companyResponseEntity">Holds company response detail entity data.</param>
        /// <returns>A task that represents company response entity data is saved or updated.</returns>
        Task<bool> UpsertConverationStateAsync(CompanyResponseEntity companyResponseEntity);

        /// <summary>
        /// Get already saved company response detail from Microsoft Azure Table storage table.
        /// </summary>
        /// <param name="responseId">Appropriate row data will be fetched based on the response id received from the bot.</param>
        /// <returns><see cref="Task"/>Already saved entity detail.</returns>
        Task<CompanyResponseEntity> GetCompanyResponseEntityAsync(string responseId);
    }
}
