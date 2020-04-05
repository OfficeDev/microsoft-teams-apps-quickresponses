// <copyright file="ConversationEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Models
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Conversation entity to store conversation id and aad object id of user for mapping purpose.
    /// </summary>
    public class ConversationEntity : TableEntity
    {
        /// <summary>
        /// Conversation table store partition key name.
        /// </summary>
        public const string ConversationPartitionKey = "ConversationEntity";

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversationEntity"/> class.
        /// </summary>
        public ConversationEntity()
        {
            this.PartitionKey = ConversationPartitionKey;
        }

        /// <summary>
        /// Gets or sets conversation id.
        /// </summary>
        public string ConversationId { get; set; }

        /// <summary>
        /// Gets or sets Aad object id of user.
        /// </summary>
        public string UserId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }
    }
}
