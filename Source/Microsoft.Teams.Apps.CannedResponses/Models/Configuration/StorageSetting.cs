// <copyright file="StorageSetting.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Models
{
    /// <summary>
    /// Class which will help to provide Microsoft Azure Table storage settings for Canned Responses app.
    /// </summary>
    public class StorageSetting : BotSetting
    {
        /// <summary>
        /// Gets or sets Microsoft Azure Table storage connection string.
        /// </summary>
        public string ConnectionString { get; set; }
    }
}
