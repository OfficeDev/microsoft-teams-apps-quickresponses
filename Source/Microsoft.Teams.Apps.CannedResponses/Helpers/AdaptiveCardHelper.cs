// <copyright file="AdaptiveCardHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Helpers
{
    using System.Collections.Generic;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;

    /// <summary>
    /// Adaptive card helper class to refresh or send notification cards.
    /// </summary>
    public static class AdaptiveCardHelper
    {
        /// <summary>
        /// Refresh the card for approved or rejected request.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="activityId">Activity id of the posted card for refreshing the card.</param>
        /// <param name="attachment">Attachment card for approved or rejected request response.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        internal static async Task RefreshCardAsync(
             ITurnContext<IMessageActivity> turnContext,
             string activityId,
             Attachment attachment)
        {
            var updateCardActivity = new Activity(ActivityTypes.Message)
            {
                Id = activityId,
                Conversation = turnContext.Activity.Conversation,
                Attachments = new List<Attachment> { attachment },
            };

            // Send approved or rejected card response to refresh the same card.
            await turnContext.UpdateActivityAsync(updateCardActivity, cancellationToken: default);
        }

        /// <summary>
        /// Send notification card for approved or rejected request.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="userNotification">Contains user notification card.</param>
        /// <param name="conversationId">Contains user conversation id.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        internal static async Task SendNotificationCardAsync(
             ITurnContext<IMessageActivity> turnContext,
             IMessageActivity userNotification,
             string conversationId,
             CancellationToken cancellationToken)
        {
            if (userNotification != null)
            {
                userNotification.Conversation = new ConversationAccount { Id = conversationId };
                await turnContext.Adapter.SendActivitiesAsync(turnContext, new Activity[] { (Activity)userNotification }, cancellationToken);
            }
        }

        /// <summary>
        /// Get team id from the deep link URL received.
        /// </summary>
        /// <param name="teamIdDeepLink">Deep link to get the team id.</param>
        /// <returns>A team id from the deep link URL.</returns>
        /// <remarks>
        /// Team id regex match for a pattern like See https://teams.microsoft.com/l/team/19%3a64c719819fb1412db8a28fd4a30b581a%40thread.tacv2/conversations?groupId=53b4782c-7c98-4449-993a-441870d10af9&amp;tenantId=72f988bf-86f1-41af-91ab-2d7cd011db47.
        /// Regex will get 19%3a64c719819fb1412db8a28fd4a30b581a%40thread.tacv2
        /// </remarks>
        internal static string ParseTeamIdFromDeepLink(string teamIdDeepLink)
        {
            var match = Regex.Match(teamIdDeepLink, @"teams.microsoft.com/l/team/(\S+)/");

            if (!match.Success)
            {
                return string.Empty;
            }

            return HttpUtility.UrlDecode(match.Groups[1].Value);
        }
    }
}
