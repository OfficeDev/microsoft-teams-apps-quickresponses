// <copyright file="CompanyResponseStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.Providers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CannedResponses.Common.Interfaces;
    using Microsoft.Teams.Apps.CannedResponses.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which helps in storing, updating, deleting company responses data in Microsoft Azure Table storage.
    /// </summary>
    public class CompanyResponseStorageProvider : BaseStorageProvider, ICompanyResponseStorageProvider
    {
        /// <summary>
        /// Represents company response entity name.
        /// </summary>
        private const string CompanyResponseEntity = "CompanyResponseEntity";

        /// <summary>
        /// Represents user id string.
        /// </summary>
        private const string UserId = "UserId";

        /// <summary>
        /// Initializes a new instance of the <see cref="CompanyResponseStorageProvider"/> class.
        /// Handles Microsoft Azure Table storage read write operations.
        /// </summary>
        /// <param name="options">A set of key/value application configuration properties for Microsoft Azure Table storage.</param>
        public CompanyResponseStorageProvider(IOptions<StorageSetting> options)
            : base(options?.Value.ConnectionString, CompanyResponseEntity)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }
        }

        /// <summary>
        /// Get company responses data from Microsoft Azure Table storage.
        /// </summary>
        /// <returns>A task that holds company response entity data in collection.</returns>
        public async Task<IEnumerable<CompanyResponseEntity>> GetCompanyResponsesDataAsync()
        {
            await this.EnsureInitializedAsync();

            string userIdCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, Constants.CompanyResponseEntityPartitionKey);
            TableQuery<CompanyResponseEntity> query = new TableQuery<CompanyResponseEntity>().Where(userIdCondition);
            TableContinuationToken continuationToken = null;
            var companyResponseCollection = new List<CompanyResponseEntity>();

            do
            {
                var queryResult = await this.ResponsesCloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                if (queryResult?.Results != null)
                {
                    companyResponseCollection.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return companyResponseCollection;
        }

        /// <summary>
        /// Get company responses from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="userId">Id of the user to fetch the responses submitted by him.</param>
        /// <returns>A task that represent collection to hold company responses data.</returns>
        public async Task<IEnumerable<CompanyResponseEntity>> GetUserCompanyResponseAsync(string userId)
        {
            await this.EnsureInitializedAsync();

            string partitionKeyCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, Constants.CompanyResponseEntityPartitionKey);
            string userIdCondition = TableQuery.GenerateFilterCondition(UserId, QueryComparisons.Equal, userId);
            string combinedFilterCondition = TableQuery.CombineFilters(partitionKeyCondition, TableOperators.And, userIdCondition);

            TableQuery<CompanyResponseEntity> query = new TableQuery<CompanyResponseEntity>().Where(combinedFilterCondition);
            TableContinuationToken continuationToken = null;
            var userRequestCollection = new List<CompanyResponseEntity>();

            do
            {
                var queryResult = await this.ResponsesCloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                if (queryResult?.Results != null)
                {
                    userRequestCollection.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return userRequestCollection?.OrderByDescending(request => request.LastUpdatedDate);
        }

        /// <summary>
        /// Delete company response details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="entity">Holds company response detail entity data.</param>
        /// <returns>A task that represents company response entity data is deleted.</returns>
        public async Task<bool> DeleteEntityAsync(CompanyResponseEntity entity)
        {
            await this.EnsureInitializedAsync();
            TableOperation deleteOperation = TableOperation.Delete(entity);
            var result = await this.ResponsesCloudTable.ExecuteAsync(deleteOperation);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Stores or update company response data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="companyResponseEntity">Holds company response detail entity data.</param>
        /// <returns>A task that represents company response entity data is saved or updated.</returns>
        public async Task<bool> UpsertConverationStateAsync(CompanyResponseEntity companyResponseEntity)
        {
            var result = await this.StoreOrUpdateEntityAsync(companyResponseEntity);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Get already saved company response detail from Microsoft Azure Table storage table.
        /// </summary>
        /// <param name="responseId">Appropriate row data will be fetched based on the response id received from the bot.</param>
        /// <returns><see cref="Task"/>Already saved entity detail.</returns>
        public async Task<CompanyResponseEntity> GetCompanyResponseEntityAsync(string responseId)
        {
            await this.EnsureInitializedAsync();

            // "When there is no company response created and messaging extension is open by Admin, table initialization is required
            // before creating search index or data-source or indexer." In this case response id will be null.
            if (string.IsNullOrEmpty(responseId))
            {
                return null;
            }

            var searchOperation = TableOperation.Retrieve<CompanyResponseEntity>(Constants.CompanyResponseEntityPartitionKey, responseId);
            var searchResult = await this.ResponsesCloudTable.ExecuteAsync(searchOperation);

            return (CompanyResponseEntity)searchResult.Result;
        }

        /// <summary>
        /// Stores or update company response details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="entity">Holds company response detail entity data.</param>
        /// <returns>A task that represents a company response data that is saved or updated.</returns>
        private async Task<TableResult> StoreOrUpdateEntityAsync(CompanyResponseEntity entity)
        {
            try
            {
                await this.EnsureInitializedAsync();
                TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(entity);
                return await this.ResponsesCloudTable.ExecuteAsync(addOrUpdateOperation);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
