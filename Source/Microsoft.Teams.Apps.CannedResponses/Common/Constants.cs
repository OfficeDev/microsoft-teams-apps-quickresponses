// <copyright file="Constants.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common
{
    /// <summary>
    /// Class that holds application constants that are used in multiple files.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Partition key for company response entity table.
        /// </summary>
        public const string CompanyResponseEntityPartitionKey = "CompanyResponseEntity";

        /// <summary>
        /// Your responses command id in the manifest file.
        /// </summary>
        public const string YourResponseCommandId = "yourResponses";

        /// <summary>
        ///  Company responses command id in the manifest file.
        /// </summary>
        public const string CompanyResponseCommandId = "companyResponses";

        /// <summary>
        /// Command for approve the suggested response request.
        /// </summary>
        public const string ApproveCommand = "approve";

        /// <summary>
        /// Command for reject the suggested response request.
        /// </summary>
        public const string RejectCommand = "reject";

        /// <summary>
        /// Date time format to support adaptive card text feature.
        /// </summary>
        /// <remarks>
        /// refer adaptive card text feature https://docs.microsoft.com/en-us/adaptive-cards/authoring-cards/text-features#datetime-formatting-and-localization
        /// </remarks>
        public const string Rfc3339DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'";

        /// <summary>
        /// Azure Search service maximum search result count for company response entity.
        /// </summary>
        public const int DefaultSearchResultCount = 200;
    }
}
