// <copyright file="BotSetting.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Models
{
    /// <summary>
    /// Class which will help to provide Bot settings for Canned Responses application.
    /// </summary>
    public class BotSetting
    {
        /// <summary>
        /// Gets or sets application base Uri which helps in generating Customer Token.
        /// </summary>
        public string AppBaseUri { get; set; }

        /// <summary>
        /// Gets or sets security key which helps in generating Customer Token.
        /// </summary>
        public string SecurityKey { get; set; }

        /// <summary>
        /// Gets or sets the deep link of the team channel.
        /// </summary>
        public string TeamIdDeepLink { get; set; }

        /// <summary>
        /// Gets or sets application tenant id.
        /// </summary>
        public string TenantId { get; set; }
    }
}
