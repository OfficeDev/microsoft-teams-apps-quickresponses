// <copyright file="AdaptiveSubmitActionData.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Models
{
    using Microsoft.Bot.Schema;
    using Newtonsoft.Json;

    /// <summary>
    /// Adaptive Card Action class to post question data.
    /// </summary>
    public class AdaptiveSubmitActionData
    {
        /// <summary>
        /// Gets or sets the Teams-specific action.
        /// </summary>
        [JsonProperty("adaptivecardactions")]
        public CardAction AdaptiveCardActions { get; set; }

        /// <summary>
        /// Gets or sets the remark value entered by Admin who rejected the request.
        /// </summary>
        [JsonProperty("approvalremark")]
        public string ApprovalRemark { get; set; }

        /// <summary>
        ///  Gets or sets the updated Label of a question while approving the request.
        /// </summary>
        [JsonProperty("updatedquestioncategory")]
        public string UpdatedQuestionCategory { get; set; }

        /// <summary>
        /// Gets or sets the updated question text while approving the request.
        /// </summary>
        [JsonProperty("updatedquestiontext")]
        public string UpdatedQuestionText { get; set; }

        /// <summary>
        /// Gets or sets the updated question text while approving the request.
        /// </summary>
        [JsonProperty("updatedresponsetext")]
        public string UpdatedResponseText { get; set; }

        /// <summary>
        /// Gets or sets Response Id.
        /// </summary>
        public string ResponseId { get; set; }

        /// <summary>
        /// Gets or sets approval status of suggested response like Pending, Approved or Rejected.
        /// </summary>
        public string ApprovalStatus { get; set; }
    }
}
