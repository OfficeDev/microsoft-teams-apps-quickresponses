// <copyright file="IUserStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.Interfaces
{
    using System.Threading.Tasks;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// Interface for storage helper which helps in storing/updating user responses data in Microsoft Azure Table storage.
    /// </summary>
    public interface IUserStorageHelper
    {
        /// <summary>
        /// Store user request details to Microsoft Azure Table storage.
        /// </summary>
        /// <param name="aadObjectId">Represents Azure active directory object id of user for current turn of bot.</param>
        /// <param name="userRequestDetails">User new request detail.</param>
        /// <returns>Represent a task queued for operation.</returns>
        Task<bool> AddNewUserRequestDetailsAsync(string aadObjectId, AddUserResponseRequestDetail userRequestDetails);

        /// <summary>
        /// Update user request details to Microsoft Azure Table storage.
        /// </summary>
        /// <param name="activity">Represents activity for current turn of bot.</param>
        /// <param name="userRequestDetails">User new request detail.</param>
        /// <returns>Represent a task queued for operation.</returns>
        Task<bool> UpdateUserRequestDetailsAsync(IInvokeActivity activity, AddUserResponseRequestDetail userRequestDetails);
    }
}
