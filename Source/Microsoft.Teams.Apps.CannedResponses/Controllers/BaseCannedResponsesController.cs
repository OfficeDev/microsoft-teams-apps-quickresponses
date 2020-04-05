// <copyright file="BaseCannedResponsesController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Controllers
{
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CannedResponses.Models;
    using Error = Microsoft.Teams.Apps.CannedResponses.Models.ErrorResponse;

    /// <summary>
    /// Base controller to handle user and company response API operations.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class BaseCannedResponsesController : ControllerBase
    {
        /// <summary>
        /// Instance of application insights telemetry client.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseCannedResponsesController"/> class.
        /// </summary>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        public BaseCannedResponsesController(TelemetryClient telemetryClient)
        {
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Records event data to Application Insights telemetry client.
        /// </summary>
        /// <param name="eventName">Name of the event.</param>
        /// <param name="userId">Azure active directory object id of the user.</param>
        public void RecordEvent(string eventName, string userId)
        {
            this.telemetryClient.TrackEvent(eventName, new Dictionary<string, string>
            {
                { "userId", userId },
            });
        }

        /// <summary>
        /// Creates the error response as per the status codes.
        /// </summary>
        /// <param name="statusCode">Describes the type of error.</param>
        /// <param name="errorMessage">Describes the error message.</param>
        /// <returns>Returns error response with appropriate message and status code.</returns>
        protected IActionResult GetErrorResponse(int statusCode, string errorMessage)
        {
            switch (statusCode)
            {
                case StatusCodes.Status400BadRequest:
                    return this.StatusCode(
                      StatusCodes.Status400BadRequest,
                      new Error
                      {
                          StatusCode = "badRequest",
                          ErrorMessage = errorMessage,
                      });
                default:
                    return this.StatusCode(
                      StatusCodes.Status500InternalServerError,
                      new Error
                      {
                          StatusCode = "internalServerError",
                          ErrorMessage = errorMessage,
                      });
            }
        }

        /// <summary>
        /// Get claims of user.
        /// </summary>
        /// <returns>User claims.</returns>
        protected JwtClaims GetUserClaims()
        {
            var claims = this.User.Claims;
            var jwtClaims = new JwtClaims
            {
                FromId = claims.Where(claim => claim.Type == "userObjectIdentifier").Select(claim => claim.Value).First(),
                ServiceUrl = claims.Where(claim => claim.Type == "serviceURL").Select(claim => claim.Value).First(),
            };

            return jwtClaims;
        }
    }
}