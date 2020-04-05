// <copyright file="CompanyStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Helpers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CannedResponses.Common.Interfaces;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// Implements storage helper which stores company responses data in Microsoft Azure Table storage.
    /// </summary>
    public class CompanyStorageHelper : ICompanyStorageHelper
    {
        /// <summary>
        /// Sets approval status as pending whenever new suggestion is submitted.
        /// </summary>
        private const string PendingRequestStatus = "Pending";

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<CompanyStorageHelper> logger;

        /// <summary>
        /// Storage provider for working with company responses data in Microsoft Azure Table storage.
        /// </summary>
        private readonly ICompanyResponseStorageProvider companyResponseStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="CompanyStorageHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="companyResponseStorageProvider">Company response storage provider dependency injection.</param>
        public CompanyStorageHelper(
            ILogger<CompanyStorageHelper> logger,
            ICompanyResponseStorageProvider companyResponseStorageProvider)
        {
            this.logger = logger;
            this.companyResponseStorageProvider = companyResponseStorageProvider;
        }

        /// <summary>
        /// Store user suggestion to Microsoft Azure Table storage.
        /// </summary>
        /// <param name="activity">Represents activity for current turn of bot.</param>
        /// <param name="userSuggestionDetails">New suggestion detail.</param>
        /// <returns>Represent a task queued for operation.</returns>
        public async Task<CompanyResponseEntity> AddNewSuggestionAsync(IInvokeActivity activity, AddUserResponseRequestDetail userSuggestionDetails)
        {
            userSuggestionDetails = userSuggestionDetails ?? throw new ArgumentNullException(nameof(userSuggestionDetails));
            activity = activity ?? throw new ArgumentNullException(nameof(activity));

            var userResponse = new CompanyResponseEntity()
            {
                QuestionLabel = userSuggestionDetails.Label,
                QuestionText = userSuggestionDetails.Question,
                ResponseText = userSuggestionDetails.Response,
                ResponseId = Guid.NewGuid().ToString(),
                UserId = activity.From.AadObjectId,
                LastUpdatedDate = DateTime.UtcNow,
                CreatedDate = DateTime.UtcNow,
                ApprovalStatus = PendingRequestStatus,
                CreatedBy = activity.From.Name,
                ApprovedOrRejectedDate = DateTime.UtcNow,
                UserPrincipalName = userSuggestionDetails.UPN,
            };

            await this.companyResponseStorageProvider.UpsertConverationStateAsync(userResponse);
            return userResponse;
        }

        /// <summary>
        /// Store user rejected data to Microsoft Azure Table storage.
        /// </summary>
        /// <param name="cardPostedData">Represents card submitted data.</param>
        /// <param name="name">Gets or sets display friendly name.</param>
        /// <param name="aadObjectId">Gets or sets this account's object ID within Azure Active Directory (AAD).</param>
        /// <returns>Represent a task queued for operation.</returns>
        public CompanyResponseEntity AddRejectedData(AdaptiveSubmitActionData cardPostedData, string name, string aadObjectId)
        {
            cardPostedData = cardPostedData ?? throw new ArgumentNullException(nameof(cardPostedData));

            CompanyResponseEntity companyResponseEntity;
            companyResponseEntity = this.companyResponseStorageProvider.GetCompanyResponseEntityAsync(cardPostedData.ResponseId).GetAwaiter().GetResult();
            companyResponseEntity.ApprovalStatus = cardPostedData.ApprovalStatus;
            companyResponseEntity.ApprovedOrRejectedBy = name;
            companyResponseEntity.ApproverUserId = aadObjectId;
            companyResponseEntity.ApprovedOrRejectedDate = DateTime.UtcNow;
            companyResponseEntity.ApprovalRemark = cardPostedData.ApprovalRemark;

            return companyResponseEntity;
        }

        /// <summary>
        /// Store user approved data to Microsoft Azure Table storage.
        /// </summary>
        /// <param name="cardPostedData">Represents card submitted data.</param>
        /// <param name="name">Gets or sets display friendly name.</param>
        /// <param name="aadObjectId">Gets or sets this account's object ID within Azure Active Directory (AAD).</param>
        /// <returns>Represent a task queued for operation.</returns>
        public CompanyResponseEntity AddApprovedData(AdaptiveSubmitActionData cardPostedData, string name, string aadObjectId)
        {
            cardPostedData = cardPostedData ?? throw new ArgumentNullException(nameof(cardPostedData));

            CompanyResponseEntity companyResponseEntity;
            companyResponseEntity = this.companyResponseStorageProvider.GetCompanyResponseEntityAsync(cardPostedData.ResponseId).GetAwaiter().GetResult();
            companyResponseEntity.QuestionLabel = cardPostedData.UpdatedQuestionCategory;
            companyResponseEntity.QuestionText = cardPostedData.UpdatedQuestionText;
            companyResponseEntity.ResponseText = cardPostedData.UpdatedResponseText;
            companyResponseEntity.ApprovalStatus = cardPostedData.ApprovalStatus;
            companyResponseEntity.ApprovedOrRejectedBy = name;
            companyResponseEntity.ApproverUserId = aadObjectId;
            companyResponseEntity.ApprovedOrRejectedDate = DateTime.UtcNow;

            return companyResponseEntity;
        }
    }
}
