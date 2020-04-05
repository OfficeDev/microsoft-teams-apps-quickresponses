// <copyright file="ICompanyStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.Interfaces
{
    using System.Threading.Tasks;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// Interface for storage helper class which helps in storing company responses data in Microsoft Azure Table storage.
    /// </summary>
    public interface ICompanyStorageHelper
    {
        /// <summary>
        /// Store user suggestion to Microsoft Azure Table storage.
        /// </summary>
        /// <param name="activity">Represents activity for current turn of bot.</param>
        /// <param name="userSuggestionDetails">New suggestion detail.</param>
        /// <returns>Represent a task queued for operation.</returns>
        Task<CompanyResponseEntity> AddNewSuggestionAsync(IInvokeActivity activity, AddUserResponseRequestDetail userSuggestionDetails);

        /// <summary>
        /// Store user rejected data to Microsoft Azure Table storage.
        /// </summary>
        /// <param name="cardPostedData">Represents card submitted data.</param>
        /// <param name="name">Gets or sets display friendly name.</param>
        /// <param name="aadObjectId">Gets or sets this account's object ID within Azure Active Directory (AAD).</param>
        /// <returns>Represent a task queued for operation.</returns>
        CompanyResponseEntity AddRejectedData(AdaptiveSubmitActionData cardPostedData, string name, string aadObjectId);

        /// <summary>
        /// Store user approved data to Microsoft Azure Table storage.
        /// </summary>
        /// <param name="cardPostedData">Represents card submitted data.</param>
        /// <param name="name">Gets or sets display friendly name.</param>
        /// <param name="aadObjectId">Gets or sets this account's object ID within Azure Active Directory (AAD).</param>
        /// <returns>Represent a task queued for operation.</returns>
        CompanyResponseEntity AddApprovedData(AdaptiveSubmitActionData cardPostedData, string name, string aadObjectId);
    }
}
