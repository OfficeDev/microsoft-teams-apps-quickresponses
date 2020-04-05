// <copyright file="AddUserResponseRequestDetail.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// User response new request details class used to send new request data.
    /// </summary>
    public class AddUserResponseRequestDetail
    {
        /// <summary>
        /// Gets or sets question label text value.
        /// </summary>
        [JsonProperty("label")]
        public string Label { get; set; }

        /// <summary>
        /// Gets or sets response id of a request.
        /// </summary>
        [JsonProperty("responseid")]
        public string ResponseId { get; set; }

        /// <summary>
        /// Gets or sets question text value.
        /// </summary>
        [JsonProperty("question")]
        public string Question { get; set; }

        /// <summary>
        /// Gets or sets response text value..
        /// </summary>
        [JsonProperty("response")]
        public string Response { get; set; }

        /// <summary>
        /// Gets or sets command context text value.
        /// </summary>
        [JsonProperty("commandContext")]
        public string CommandContext { get; set; }

        /// <summary>
        /// Gets or sets user principal name(i.e. email address) of the user.
        /// </summary>
        [JsonProperty("UPN")]
        public string UPN { get; set; }
    }
}
