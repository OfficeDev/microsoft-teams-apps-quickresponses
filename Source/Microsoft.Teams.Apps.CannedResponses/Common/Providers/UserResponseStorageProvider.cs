// <copyright file="UserResponseStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CannedResponses.Common.Interfaces;
    using Microsoft.Teams.Apps.CannedResponses.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which stores user responses data in Microsoft Azure Table storage.
    /// </summary>
    public class UserResponseStorageProvider : BaseStorageProvider, IUserResponseStorageProvider
    {
        /// <summary>
        /// Represents user response entity name.
        /// </summary>
        private const string UserResponseEntity = "UserResponseEntity";

        /// <summary>
        /// Represents row key string.
        /// </summary>
        private const string RowKey = "RowKey";

        /// <summary>
        /// Initializes a new instance of the <see cref="UserResponseStorageProvider"/> class.
        /// Handles Microsoft Azure Table storage read write operations.
        /// </summary>
        /// <param name="options">A set of key/value application configuration properties for Microsoft Azure Table storage.</param>
        public UserResponseStorageProvider(IOptions<StorageSetting> options)
            : base(options?.Value.ConnectionString, UserResponseEntity)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }
        }

        /// <summary>
        /// Get user responses data from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="userId">User Id for which need to fetch data.</param>
        /// <returns>A task that represent collection to hold user responses data.</returns>
        public async Task<IEnumerable<UserResponseEntity>> GetUserResponsesDataAsync(string userId)
        {
            await this.EnsureInitializedAsync();
            if (string.IsNullOrEmpty(userId))
            {
                return null;
            }

            string userIdCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, userId);
            TableQuery<UserResponseEntity> query = new TableQuery<UserResponseEntity>().Where(userIdCondition);
            TableContinuationToken continuationToken = null;
            var userResponseCollection = new List<UserResponseEntity>();

            do
            {
                var queryResult = await this.ResponsesCloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                if (queryResult?.Results != null)
                {
                    userResponseCollection.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return userResponseCollection;
        }

        /// <summary>
        /// Get user responses data from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="responseId">Response Id for which need to fetch data.</param>
        /// <returns>A task that represent collection to hold user responses data.</returns>
        public async Task<IEnumerable<UserResponseEntity>> GetUserResponseDataAsync(string responseId)
        {
            await this.EnsureInitializedAsync();
            if (string.IsNullOrEmpty(responseId))
            {
                return null;
            }

            string responseIdCondition = TableQuery.GenerateFilterCondition(RowKey, QueryComparisons.Equal, responseId);
            TableQuery<UserResponseEntity> query = new TableQuery<UserResponseEntity>().Where(responseIdCondition);
            TableContinuationToken continuationToken = null;
            var userResponseCollection = new List<UserResponseEntity>();

            do
            {
                var queryResult = await this.ResponsesCloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                if (queryResult?.Results != null)
                {
                    userResponseCollection.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return userResponseCollection;
        }

        /// <summary>
        /// Delete user response details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="userResponseIds">Holds user response Id data.</param>
        /// <returns>A task that represents user response entity data is saved or updated.</returns>
        public async Task<bool> DeleteEntityAsync(IEnumerable<string> userResponseIds)
        {
            if (userResponseIds == null)
            {
                throw new ArgumentNullException(nameof(userResponseIds));
            }

            await this.EnsureInitializedAsync();
            var entity = new UserResponseEntity();

            foreach (var userResponseId in userResponseIds)
            {
                string responseIdCondition = TableQuery.GenerateFilterCondition(RowKey, QueryComparisons.Equal, userResponseId);
                TableQuery<UserResponseEntity> query = new TableQuery<UserResponseEntity>().Where(responseIdCondition);
                var queryResult = await this.ResponsesCloudTable.ExecuteQuerySegmentedAsync(query, null);
                entity = queryResult?.Results[0];
                TableOperation deleteOperation = TableOperation.Delete(entity);
                var result = await this.ResponsesCloudTable.ExecuteAsync(deleteOperation);
            }

            return true;
        }

        /// <summary>
        /// Stores or update user response details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="userResponseEntity">Holds user response detail entity data.</param>
        /// <returns>A task that represents user response entity data is saved or updated.</returns>
        public async Task<bool> UpsertUserResponseAsync(UserResponseEntity userResponseEntity)
        {
            var result = await this.StoreOrUpdateEntityAsync(userResponseEntity);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Stores or update user response details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="entity">Holds user response detail entity data.</param>
        /// <returns>A task that represents user response entity data is saved or updated.</returns>
        private async Task<TableResult> StoreOrUpdateEntityAsync(UserResponseEntity entity)
        {
            await this.EnsureInitializedAsync();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(entity);
            return await this.ResponsesCloudTable.ExecuteAsync(addOrUpdateOperation);
        }
    }
}