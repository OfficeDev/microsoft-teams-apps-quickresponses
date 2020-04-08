// <copyright file="UserResponseEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using Microsoft.Azure.Search;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Class which represents use response entity model.
    /// </summary>
    public class UserResponseEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets User Id.
        /// </summary>
        [IsFilterable]
        public string UserId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
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
        /// Gets or sets response updated date and time.
        /// </summary>
        [IsSortable]
        public DateTime LastUpdatedDate { get; set; }
    }
}
