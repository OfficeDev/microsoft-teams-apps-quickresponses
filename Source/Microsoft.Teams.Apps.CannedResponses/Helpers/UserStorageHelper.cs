// <copyright file="UserStorageHelper.cs" company="Microsoft">
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
    /// Implements user storage helper which is responsible for storing or updating data data in Microsoft Azure Table storage.
    /// </summary>
    public class UserStorageHelper : IUserStorageHelper
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<UserStorageHelper> logger;

        /// <summary>
        /// Storage provider for working with company responses data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IUserResponseStorageProvider userResponseStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserStorageHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="userResponseStorageProvider">Company response storage provider dependency injection.</param>
        public UserStorageHelper(
            ILogger<UserStorageHelper> logger,
            IUserResponseStorageProvider userResponseStorageProvider)
        {
            this.logger = logger;
            this.userResponseStorageProvider = userResponseStorageProvider;
        }

        /// <summary>
        /// Store user request details to Microsoft Azure Table storage.
        /// </summary>
        /// <param name="aadObjectId">Represents Azure active directory object id of user for current turn of bot.</param>
        /// <param name="userRequestDetails">User new request detail.</param>
        /// <returns>Represent a task queued for operation.</returns>
        public async Task<bool> AddNewUserRequestDetailsAsync(string aadObjectId, AddUserResponseRequestDetail userRequestDetails)
        {
            if (userRequestDetails != null)
            {
                var userResponse = new UserResponseEntity()
                {
                    QuestionLabel = userRequestDetails.Label,
                    QuestionText = userRequestDetails.Question,
                    ResponseText = userRequestDetails.Response,
                    ResponseId = Guid.NewGuid().ToString(),
                    UserId = aadObjectId,
                    LastUpdatedDate = DateTime.UtcNow,
                };

                return await this.userResponseStorageProvider.UpsertUserResponseAsync(userResponse);
            }

            return false;
        }

        /// <summary>
        /// Update user request details to Microsoft Azure Table storage.
        /// </summary>
        /// <param name="activity">Represents activity for current turn of bot.</param>
        /// <param name="userRequestDetails">User new request detail.</param>
        /// <returns>Represent a task queued for operation.</returns>
        public async Task<bool> UpdateUserRequestDetailsAsync(IInvokeActivity activity, AddUserResponseRequestDetail userRequestDetails)
        {
            if (userRequestDetails != null && activity != null)
            {
                var userResponse = new UserResponseEntity()
                {
                    QuestionLabel = userRequestDetails.Label,
                    QuestionText = userRequestDetails.Question,
                    ResponseText = userRequestDetails.Response,
                    UserId = activity.From.AadObjectId,
                    LastUpdatedDate = DateTime.UtcNow,
                    ResponseId = userRequestDetails.ResponseId,
                };

                return await this.userResponseStorageProvider.UpsertUserResponseAsync(userResponse);
            }

            return false;
        }
    }
}
