// <copyright file="CannedResponsesActivityHandlerOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Bot
{
    /// <summary>
    /// The CannedResponsesActivityHandlerOptions are the options for the <see cref="CannedResponsesActivityHandler" /> bot.
    /// </summary>
    public sealed class CannedResponsesActivityHandlerOptions
    {
        /// <summary>
        /// Gets or sets a value indicating whether the response to a message should be in all uppercase.
        /// </summary>
        public bool UpperCaseResponse { get; set; }

        /// <summary>
        /// Gets or sets application base URL used to return success or failure task module result.
        /// </summary>
        public string AppBaseUri { get; set; }
    }
}