// <copyright file="CompanyResponseEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using Microsoft.Azure.Search;
    using Microsoft.Teams.Apps.CannedResponses.Common;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Class which represents company response entity model.
    /// </summary>
    public class CompanyResponseEntity : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CompanyResponseEntity"/> class.
        /// Holds company responses details data.
        /// </summary>
        public CompanyResponseEntity()
        {
            this.PartitionKey = Constants.CompanyResponseEntityPartitionKey;
        }

        /// <summary>
        /// Gets or sets Response Id.
        /// </summary>
        [Key]
        public string ResponseId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets Label of a question.
        /// </summary>
        [IsSearchable]
        public string QuestionLabel { get; set; }

        /// <summary>
        /// Gets or sets question text.
        /// </summary>
        [IsSearchable]
        public string QuestionText { get; set; }

        /// <summary>
        /// Gets or sets response text.
        /// </summary>
        [IsSearchable]
        public string ResponseText { get; set; }

        /// <summary>
        /// Gets or sets user id who suggests the response.
        /// </summary>
        public string UserId { get; set; }

        /// <summary>
        /// Gets or sets name of user who suggests the response.
        /// </summary>
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets response created date and time.
        /// </summary>
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// Gets or sets the user requested type for response like New/Edit.
        /// </summary>
        public string UserRequestType { get; set; }

        /// <summary>
        /// Gets or sets response updated date and time.
        /// </summary>
        [IsSortable]
        public DateTime LastUpdatedDate { get; set; }

        /// <summary>
        /// Gets or sets use who has updated the response.
        /// </summary>
        public string LastUpdatedBy { get; set; }

        /// <summary>
        /// Gets or sets user id who approved or rejected the suggested response.
        /// </summary>
        public string ApproverUserId { get; set; }

        /// <summary>
        /// Gets or sets name of user who approved or rejected the response for question.
        /// </summary>
        public string ApprovedOrRejectedBy { get; set; }

        /// <summary>
        /// Gets or sets approval status of suggested response like Pending, Approved or Rejected.
        /// </summary>
        [IsFilterable]
        public string ApprovalStatus { get; set; }

        /// <summary>
        /// Gets or sets the remark value entered by Admin who rejected the request.
        /// </summary>
        public string ApprovalRemark { get; set; }

        /// <summary>
        /// Gets or sets activity id of a user response request adaptive card.
        /// </summary>
        public string ActivityId { get; set; }

        /// <summary>
        /// Gets or sets datetime when the request is approved or rejected.
        /// </summary>
        [IsSortable]
        public DateTime ApprovedOrRejectedDate { get; set; }

        /// <summary>
        /// Gets or sets user principal name of the user.
        /// </summary>
        public string UserPrincipalName { get; set; }
    }
}
