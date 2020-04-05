// <copyright file="IMessagingExtensionHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.Interfaces
{
    using System.Threading.Tasks;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;

    /// <summary>
    /// Interface that handles the search activities for messaging extension.
    /// </summary>
    public interface IMessagingExtensionHelper
    {
        /// <summary>
        /// Get the results from Azure Search service and populate the result (card + preview).
        /// </summary>
        /// <param name="query">Query which the user had typed in messaging extension search field.</param>
        /// <param name="commandId">Command id to determine which tab in messaging extension has been invoked.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <param name="userObjectId">AAd object id of the user.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns><see cref="Task"/>Returns messaging extension result object, which will be used for providing the card.</returns>
        Task<MessagingExtensionResult> GetSearchResultAsync(
            string query,
            string commandId,
            int? count,
            int? skip,
            string userObjectId,
            IStringLocalizer<Strings> localizer);

        /// <summary>
        /// Get the value of the searchText parameter in the messaging extension query.
        /// </summary>
        /// <param name="query">Contains messaging extension query keywords.</param>
        /// <returns>A value of the searchText parameter.</returns>
        string GetSearchQueryString(MessagingExtensionQuery query);
    }
}
