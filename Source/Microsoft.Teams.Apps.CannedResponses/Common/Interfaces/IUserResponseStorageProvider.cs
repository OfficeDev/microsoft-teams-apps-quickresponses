// <copyright file="IUserResponseStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// Interface for provider which helps in storing/updating user response details in Microsoft Azure Table storage.
    /// </summary>
    public interface IUserResponseStorageProvider
    {
        /// <summary>
        /// Get user responses data from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="userId">Fetches user response data based on user id.</param>
        /// <returns>A task that holds user response entity data in collection.</returns>
        Task<IEnumerable<UserResponseEntity>> GetUserResponsesDataAsync(string userId);

        /// <summary>
        /// Stores or update user response data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="userResponseEntity">Holds user response detail entity data.</param>
        /// <returns>A task that represents user response entity data is saved or updated.</returns>
        Task<bool> UpsertUserResponseAsync(UserResponseEntity userResponseEntity);

        /// <summary>
        /// Delete user responses data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="userResponseIds">Holds user responses id in a collection.</param>
        /// <returns>A task that represents user response entity data is deleted.</returns>
        Task<bool> DeleteEntityAsync(IEnumerable<string> userResponseIds);

        /// <summary>
        /// Get user responses data from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="responseId">Response Id for which need to fetch data.</param>
        /// <returns>A task that represents a collection to hold user responses data.</returns>
        Task<IEnumerable<UserResponseEntity>> GetUserResponseDataAsync(string responseId);
    }
}
