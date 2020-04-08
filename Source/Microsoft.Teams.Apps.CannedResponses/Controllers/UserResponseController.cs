// <copyright file="UserResponseController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CannedResponses.Common.Interfaces;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// Controller to handle user response API operations.
    /// </summary>
    [Route("api/userresponse")]
    [ApiController]
    [Authorize]
    public class UserResponseController : BaseCannedResponsesController
    {
        /// <summary>
        /// Event name for user response HTTP get call.
        /// </summary>
        private const string RecordUserHTTPGetCall = "User responses - HTTP Get call succeeded";

        /// <summary>
        /// Event name for user response HTTP post call.
        /// </summary>
        private const string RecordUserHTTPPostCall = "User responses - HTTP Post call succeeded";

        /// <summary>
        /// Event name for user response HTTP put call.
        /// </summary>
        private const string RecordUserHTTPPutCall = "User responses - HTTP Put call succeeded";

        /// <summary>
        /// Event name for user response HTTP delete call.
        /// </summary>
        private const string RecordUserHTTPDeleteCall = "User responses - HTTP Delete call succeeded";

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Instance of user response storage provider to update response and get information of responses.
        /// </summary>
        private readonly IUserResponseStorageProvider userResponseStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserResponseController"/> class.
        /// </summary>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="userResponseStorageProvider">User response storage provider dependency injection.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        public UserResponseController(
            ILogger<UserResponseController> logger,
            IUserResponseStorageProvider userResponseStorageProvider,
            TelemetryClient telemetryClient)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.userResponseStorageProvider = userResponseStorageProvider;
        }

        /// <summary>
        /// Get call to retrieve list of user responses.
        /// </summary>
        /// <returns>List of user responses</returns>
        [HttpGet]
        public async Task<IActionResult> GetAsync()
        {
            try
            {
                var claims = this.GetUserClaims();
                var userResponses = await this.userResponseStorageProvider.GetUserResponsesDataAsync(claims.FromId);
                this.RecordEvent(RecordUserHTTPGetCall, claims.FromId);
                return this.Ok(userResponses?.OrderByDescending(response => response.LastUpdatedDate).Take(500));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to user response service.");
                throw;
            }
        }

        /// <summary>
        /// Post call to store user response details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="userResponseEntity">Holds user response detail entity data.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost]
        public async Task<IActionResult> PostAsync([FromBody] UserResponseEntity userResponseEntity)
        {
            try
            {
                if (string.IsNullOrEmpty(userResponseEntity?.UserId))
                {
                    this.logger.LogError("Error while creating user response details data in Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while creating user response details data in Microsoft Azure Table storage.");
                }

                var claims = this.GetUserClaims();
                this.RecordEvent(RecordUserHTTPPostCall, claims.FromId);
                return this.Ok(await this.userResponseStorageProvider.UpsertUserResponseAsync(userResponseEntity));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to user response service.");
                throw;
            }
        }

        /// <summary>
        /// Put call to update user response details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="userResponseEntity">Holds user response detail entity data.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPut]
        public async Task<IActionResult> PutAsync([FromBody] UserResponseEntity userResponseEntity)
        {
            try
            {
                if (string.IsNullOrEmpty(userResponseEntity?.UserId) || string.IsNullOrEmpty(userResponseEntity.ResponseId))
                {
                    this.logger.LogError("Error while updating user response details data in Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while updating user response details data");
                }

                var claims = this.GetUserClaims();
                this.RecordEvent(RecordUserHTTPPutCall, claims.FromId);
                return this.Ok(await this.userResponseStorageProvider.UpsertUserResponseAsync(userResponseEntity));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to user response service.");
                throw;
            }
        }

        /// <summary>
        /// Delete call to delete user response details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="userResponseIds">User selected response Ids.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost]
        [Route("deleteresponses")]
        public async Task<IActionResult> DeleteAsync([FromBody] IList<string> userResponseIds)
        {
            try
            {
                if (userResponseIds == null)
                {
                    this.logger.LogError("Error while deleting user response details data in Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while deleting user response details data in Microsoft Azure Table storage.");
                }

                var claims = this.GetUserClaims();
                this.RecordEvent(RecordUserHTTPDeleteCall, claims.FromId);
                return this.Ok(await this.userResponseStorageProvider.DeleteEntityAsync(userResponseIds));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to user response service.");
                throw;
            }
        }

        /// <summary>
        /// Get user response details data for response id from in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="responseId">User selected response Id.</param>
        /// <returns>Returns true for successful operation.</returns>
        [Route("responsedata")]
        public async Task<IActionResult> GetUserResponseDetailAsync(string responseId)
        {
            try
            {
                var userResponses = await this.userResponseStorageProvider.GetUserResponseDataAsync(responseId);

                this.logger.LogInformation("Call to user response service succeeded");
                return this.Ok(userResponses.FirstOrDefault());
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to user response service.");
                throw;
            }
        }
    }
}
