// <copyright file="ConversationStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.Providers
{
    using System;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CannedResponses.Common.Interfaces;
    using Microsoft.Teams.Apps.CannedResponses.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which helps in storing, updating user conversation data in Microsoft Azure Table storage.
    /// </summary>
    public class ConversationStorageProvider : BaseStorageProvider, IConversationStorageProvider
    {
        /// <summary>
        /// Represents conversation entity name.
        /// </summary>
        private const string ConversationEntity = "ConversationEntity";

        /// <summary>
        /// Partition key value of conversation entity table storage.
        /// </summary>
        private const string ConversationEntityParitionKey = "ConversationEntity";

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversationStorageProvider"/> class.
        /// </summary>
        /// <param name="options">A set of key/value application configuration properties.</param>
        public ConversationStorageProvider(IOptions<StorageSetting> options)
            : base(options?.Value.ConnectionString, ConversationEntity)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }
        }

        /// <summary>
        /// Get already saved user conversation detail from Microsoft Azure Table storage table.
        /// </summary>
        /// <param name="userId">Appropriate row data will be fetched based on the user id received from the bot.</param>
        /// <returns>Already saved entity detail.</returns>
        public async Task<ConversationEntity> GetConversationEntityAsync(string userId)
        {
            await this.EnsureInitializedAsync();

            if (string.IsNullOrEmpty(userId))
            {
                return null;
            }

            var searchOperation = TableOperation.Retrieve<ConversationEntity>(ConversationEntityParitionKey, userId);
            var searchResult = await this.ResponsesCloudTable.ExecuteAsync(searchOperation);

            return (ConversationEntity)searchResult.Result;
        }

        /// <summary>
        /// Add the conversation entity object in table storage.
        /// </summary>
        /// <param name="conversationEntity">Conversation table entity.</param>
        /// <returns>A <see cref="Task"/> of type bool where true represents conversation entity object is added in table storage successfully while false indicates failure in saving data.</returns>
        public async Task<bool> AddConversationEntityAsync(ConversationEntity conversationEntity)
        {
            await this.EnsureInitializedAsync();
            TableOperation insertOrMergeOperation = TableOperation.InsertOrReplace(conversationEntity);
            TableResult result = await this.ResponsesCloudTable.ExecuteAsync(insertOrMergeOperation);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }
    }
}
