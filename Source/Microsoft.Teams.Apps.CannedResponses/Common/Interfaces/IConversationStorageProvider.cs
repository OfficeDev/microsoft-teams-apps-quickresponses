// <copyright file="IConversationStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.Interfaces
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// Interface for conversation storage provider class which helps in storing user conversation details in Microsoft Azure Table storage.
    /// </summary>
    public interface IConversationStorageProvider
    {
        /// <summary>
        /// Add the conversation entity object in table storage.
        /// </summary>
        /// <param name="conversationEntity">Conversation table entity.</param>
        /// <returns>A <see cref="Task"/> of type bool where true represents conversation entity object is added in table storage successfully while false indicates failure in saving data.</returns>
        Task<bool> AddConversationEntityAsync(ConversationEntity conversationEntity);

        /// <summary>
        /// Get already saved user conversation detail from Microsoft Azure Table storage table.
        /// </summary>
        /// <param name="userId">Appropriate row data will be fetched based on the user id received from the bot.</param>
        /// <returns>Already saved entity detail.</returns>
        Task<ConversationEntity> GetConversationEntityAsync(string userId);
    }
}
