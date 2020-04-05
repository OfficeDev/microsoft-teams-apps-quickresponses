// <copyright file="CompanyResponseController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Controllers
{
    using System;
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
    /// Controller to handle company response API operations.
    /// </summary>
    [Route("api/companyresponse")]
    [ApiController]
    [Authorize]
    public class CompanyResponseController : BaseCannedResponsesController
    {
        /// <summary>
        /// Maximum retrieve count of company responses using Azure Search service.
        /// </summary>
        private const int MaximumSearchResultCount = 1000;

        /// <summary>
        /// Event name for company response get HTTP call.
        /// </summary>
        private const string RecordCompanyHTTPGetCall = "Company responses - HTTP Get call succeeded";

        /// <summary>
        /// Event name for company response get HTTP call for user.
        /// </summary>
        private const string RecordCompanyUserHTTPGetCall = "Company responses - for user HTTP Get call succeeded";

        /// <summary>
        /// Event name for company response post HTTP call.
        /// </summary>
        private const string RecordCompanyHTTPPostCall = "Company responses - HTTP Post call succeeded";

        /// <summary>
        /// Event name for company response put HTTP call.
        /// </summary>
        private const string RecordCompanyHTTPPutCall = "Company responses - HTTP Put call succeeded";

        /// <summary>
        /// Event name for company response delete HTTP call.
        /// </summary>
        private const string RecordCompanyHTTPDeleteCall = "Company responses - HTTP Delete call succeeded";

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Storage provider for working with company responses data in Microsoft Azure Table storage.
        /// </summary>
        private readonly ICompanyResponseStorageProvider companyResponseStorageProvider;

        /// <summary>
        /// Helper for working with Microsoft search service on Microsoft Azure Table storage.
        /// </summary>
        private readonly ICompanyResponseSearchService companyResponseSearchService;

        /// <summary>
        /// Initializes a new instance of the <see cref="CompanyResponseController"/> class.
        /// </summary>
        /// <param name="logger">Sends logs to the Application Insights service.</param>
        /// <param name="companyResponseStorageProvider">Company response storage provider dependency injection.</param>
        /// <param name="companyResponseSearchService">The company response search service dependency injection.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        public CompanyResponseController(
            ILogger<CompanyResponseController> logger,
            ICompanyResponseStorageProvider companyResponseStorageProvider,
            ICompanyResponseSearchService companyResponseSearchService,
            TelemetryClient telemetryClient)
            : base(telemetryClient)
        {
            this.logger = logger;
            this.companyResponseStorageProvider = companyResponseStorageProvider;
            this.companyResponseSearchService = companyResponseSearchService;
        }

        /// <summary>
        /// Get call to retrieve list of company responses.
        /// </summary>
        /// <returns>List of company responses.</returns>
        [HttpGet]
        [Route("GetCompanyResponses")]
        public async Task<IActionResult> GetAsync()
        {
            try
            {
                var companyResponses = await this.companyResponseSearchService.GetSearchCompanyResponseAsync(null, MaximumSearchResultCount, null, isTaskModuleData: true);

                var claims = this.GetUserClaims();
                this.RecordEvent(RecordCompanyHTTPGetCall, claims.FromId);
                return this.Ok(companyResponses?.Take(500));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to company response service.");
                throw;
            }
        }

        /// <summary>
        /// Get list of company responses for current user.
        /// </summary>
        /// <returns>List of company responses.</returns>
        [HttpGet]
        [Route("GetUserRequests")]
        public async Task<IActionResult> GetUserCompanyResponseAsync()
        {
            try
            {
                var claims = this.GetUserClaims();
                var userRequests = await this.companyResponseStorageProvider.GetUserCompanyResponseAsync(claims.FromId);
                this.RecordEvent(RecordCompanyUserHTTPGetCall, claims.FromId);
                this.logger.LogInformation("Call to get user requests succeeded");
                return this.Ok(userRequests?.Take(500));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to get user requests.");
                throw;
            }
        }

        /// <summary>
        /// Post call to store company response details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="companyResponseEntity">Holds company response detail entity data.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost]
        public async Task<IActionResult> PostAsync([FromBody] CompanyResponseEntity companyResponseEntity)
        {
            try
            {
                if (string.IsNullOrEmpty(companyResponseEntity?.UserId))
                {
                    this.logger.LogError("Error while creating company response details data in Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while creating company response details data in Microsoft Azure Table storage.");
                }

                var claims = this.GetUserClaims();
                this.RecordEvent(RecordCompanyHTTPPostCall, claims.FromId);
                return this.Ok(await this.companyResponseStorageProvider.UpsertConverationStateAsync(companyResponseEntity));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to company response service.");
                throw;
            }
        }

        /// <summary>
        /// Put call to update company response details data in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="companyResponseEntity">Helper for working with Microsoft Azure Table storage.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPut]
        public async Task<IActionResult> PutAsync([FromBody] CompanyResponseEntity companyResponseEntity)
        {
            try
            {
                if (string.IsNullOrEmpty(companyResponseEntity?.UserId) || string.IsNullOrEmpty(companyResponseEntity.ResponseId))
                {
                    this.logger.LogError("Error while updating company response details data in Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while updating company response details.");
                }

                var claims = this.GetUserClaims();
                this.RecordEvent(RecordCompanyHTTPPutCall, claims.FromId);
                return this.Ok(await this.companyResponseStorageProvider.UpsertConverationStateAsync(companyResponseEntity));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to company response service.");
                throw;
            }
        }

        /// <summary>
        /// Delete company response details data from Microsoft Azure Table storage.
        /// </summary>
        /// <param name="companyResponseEntity">Helper for working with Microsoft Azure Table storage.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpDelete]
        public async Task<IActionResult> DeleteAsync([FromBody] CompanyResponseEntity companyResponseEntity)
        {
            try
            {
                if (string.IsNullOrEmpty(companyResponseEntity?.UserId) || string.IsNullOrEmpty(companyResponseEntity.ResponseId))
                {
                    this.logger.LogError("Error while deleting company response details data in Microsoft Azure Table storage.");
                    return this.GetErrorResponse(StatusCodes.Status400BadRequest, "Error while deleting company response details data in Microsoft Azure Table storage.");
                }

                this.RecordEvent(RecordCompanyHTTPDeleteCall, companyResponseEntity.UserId);
                return this.Ok(await this.companyResponseStorageProvider.DeleteEntityAsync(companyResponseEntity));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while making call to company response service.");
                throw;
            }
        }
    }
}
