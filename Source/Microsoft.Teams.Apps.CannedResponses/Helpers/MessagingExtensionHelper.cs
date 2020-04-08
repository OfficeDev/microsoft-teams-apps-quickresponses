// <copyright file="MessagingExtensionHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.CannedResponses.Common;
    using Microsoft.Teams.Apps.CannedResponses.Common.Interfaces;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// Class that handles the search activities for Messaging Extension.
    /// </summary>
    public class MessagingExtensionHelper : IMessagingExtensionHelper
    {
        /// <summary>
        /// Search text parameter name in the manifest file.
        /// </summary>
        private const string SearchTextParameterName = "searchText";

        /// <summary>
        /// Helper for working with Microsoft Azure Table Search service.
        /// </summary>
        private readonly IUserResponseSearchService userResponseSearchService;

        /// <summary>
        /// Helper for working with Microsoft Azure Table Search service.
        /// </summary>
        private readonly ICompanyResponseSearchService companyResponseSearchService;

        /// <summary>
        /// Initializes a new instance of the <see cref="MessagingExtensionHelper"/> class.
        /// </summary>
        /// <param name="userResponseSearchService">The user response search service dependency injection.</param>
        /// <param name="companyResponseSearchService">The company response search service dependency injection.</param>
        public MessagingExtensionHelper(
            IUserResponseSearchService userResponseSearchService,
            ICompanyResponseSearchService companyResponseSearchService)
        {
            this.userResponseSearchService = userResponseSearchService;
            this.companyResponseSearchService = companyResponseSearchService;
        }

        /// <summary>
        /// Get the results from Azure Search service and populate the result (card + preview).
        /// </summary>
        /// <param name="query">Query which the user had typed in Messaging Extension search field.</param>
        /// <param name="commandId">Command id to determine which tab in Messaging Extension has been invoked.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <param name="userObjectId">AAd object id of the user.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns><see cref="Task"/>Returns Messaging Extension result object, which will be used for providing the card.</returns>
        public async Task<MessagingExtensionResult> GetSearchResultAsync(
            string query,
            string commandId,
            int? count,
            int? skip,
            string userObjectId,
            IStringLocalizer<Strings> localizer)
        {
            MessagingExtensionResult composeExtensionResult = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = AttachmentLayoutTypes.List,
                Attachments = new List<MessagingExtensionAttachment>(),
            };

            // commandId should be equal to Id mentioned in Manifest file under composeExtensions section.
            switch (commandId)
            {
                case Constants.YourResponseCommandId:
                    var userSearchServiceResults = await this.userResponseSearchService.SearchUserResponseAsync(query, userObjectId, count, skip);
                    composeExtensionResult = this.GetUserResponsesResult(userSearchServiceResults, localizer: localizer);
                    break;

                case Constants.CompanyResponseCommandId:
                    var companySearchServiceResults = await this.companyResponseSearchService.GetSearchCompanyResponseAsync(query, count, skip);
                    composeExtensionResult = this.GetCompanyResponsesResult(companySearchServiceResults, localizer: localizer);
                    break;
            }

            return composeExtensionResult;
        }

        /// <summary>
        /// Get the value of the searchText parameter in the Messaging Extension query.
        /// </summary>
        /// <param name="query">Contains Messaging Extension query keywords.</param>
        /// <returns>A value of the searchText parameter.</returns>
        public string GetSearchQueryString(MessagingExtensionQuery query)
        {
            var messagingExtensionInputText = query?.Parameters.FirstOrDefault(parameter => parameter.Name.Equals(SearchTextParameterName, StringComparison.OrdinalIgnoreCase));
            return messagingExtensionInputText?.Value?.ToString();
        }

        /// <summary>
        /// Get user responses result for Messaging Extension.
        /// </summary>
        /// <param name="userResponseResults">List of user search result.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns><see cref="Task"/>Returns Messaging Extension result object, which will be used for providing the card.</returns>
        private MessagingExtensionResult GetUserResponsesResult(IList<UserResponseEntity> userResponseResults, IStringLocalizer<Strings> localizer)
        {
            MessagingExtensionResult composeExtensionResult = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = AttachmentLayoutTypes.List,
                Attachments = new List<MessagingExtensionAttachment>(),
            };

            if (userResponseResults != null)
            {
                foreach (var userResponse in userResponseResults)
                {
                    var card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 2))
                    {
                        Body = new List<AdaptiveElement>
                        {
                            new AdaptiveTextBlock
                            {
                                Text = userResponse.ResponseText,
                                Wrap = true,
                            },
                        },
                    };

                    ThumbnailCard previewCard = new ThumbnailCard
                    {
                        Title = userResponse.ResponseText,
                        Text = $"{userResponse.QuestionLabel}",
                    };

                    composeExtensionResult.Attachments.Add(new Attachment
                    {
                        ContentType = AdaptiveCard.ContentType,
                        Content = card,
                    }.ToMessagingExtensionAttachment(previewCard.ToAttachment()));
                }
            }

            return composeExtensionResult;
        }

        /// <summary>
        /// Get company responses result for Messaging Extension.
        /// </summary>
        /// <param name="companySearchServiceResults">List of company search result.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns><see cref="Task"/> Returns Messaging Extension result object, which will be used for providing the card.</returns>
        private MessagingExtensionResult GetCompanyResponsesResult(IList<CompanyResponseEntity> companySearchServiceResults, IStringLocalizer<Strings> localizer)
        {
            MessagingExtensionResult composeExtensionResult = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = AttachmentLayoutTypes.List,
                Attachments = new List<MessagingExtensionAttachment>(),
            };

            if (companySearchServiceResults != null)
            {
                foreach (var companyResponse in companySearchServiceResults)
                {
                    var card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 2))
                    {
                        Body = new List<AdaptiveElement>
                        {
                            new AdaptiveTextBlock
                            {
                                Text = companyResponse.ResponseText,
                                Wrap = true,
                            },
                        },
                    };

                    ThumbnailCard previewCard = new ThumbnailCard
                    {
                        Title = companyResponse.ResponseText,
                        Text = $"{companyResponse.QuestionLabel} | {companyResponse.CreatedBy}",
                    };

                    composeExtensionResult.Attachments.Add(new Attachment
                    {
                        ContentType = AdaptiveCard.ContentType,
                        Content = card,
                    }.ToMessagingExtensionAttachment(previewCard.ToAttachment()));
                }
            }

            return composeExtensionResult;
        }
    }
}
