// <copyright file="UserCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.CannedResponses.Common;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// Class having method to return user related card attachment.
    /// </summary>
    public static class UserCard
    {
        /// <summary>
        /// Get notification card for approved request.
        /// </summary>
        /// <param name="userRequestDetails">User request details object.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>An attachment card for approved request.</returns>
        public static Attachment GetNotificationCardForApprovedRequest(CompanyResponseEntity userRequestDetails, IStringLocalizer<Strings> localizer)
        {
            if (userRequestDetails == null)
            {
                throw new ArgumentNullException(nameof(userRequestDetails));
            }

            var formattedDateTime = userRequestDetails.ApprovedOrRejectedDate.ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.InvariantCulture);
            string dateString = string.Format(CultureInfo.InvariantCulture, localizer.GetString("DateFormat"), "{{DATE(" + formattedDateTime + ", COMPACT)}}", "{{TIME(" + formattedDateTime + ")}}");

            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 2))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveContainer
                    {
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Text = localizer.GetString("NotificationCardApprovedOrRejectedTitle"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = string.Format(CultureInfo.InvariantCulture, localizer.GetString("NotificationCardStatusText"), localizer.GetString("ApprovedRequestStatusText")),
                                                Wrap = true,
                                                Color = AdaptiveTextColor.Good,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveContainer
                    {
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Text = localizer.GetString("NotificationCardLabelText"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = userRequestDetails.QuestionLabel,
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = localizer.GetString("NotificationCardQuestion"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = userRequestDetails.QuestionText,
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = localizer.GetString("NotificationCardResponse"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = userRequestDetails.ResponseText,
                                                Wrap = true,
                                            },
                                        },
                                        Style = AdaptiveContainerStyle.Emphasis,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveContainer
                    {
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Text = string.Format(CultureInfo.InvariantCulture, localizer.GetString("ApprovedCardLabelText"), dateString),
                                                Wrap = true,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                },
            };
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };

            return adaptiveCardAttachment;
        }

        /// <summary>
        /// Get notification card for rejected request.
        /// </summary>
        /// <param name="userRequestDetails">User request details object.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>An attachment.</returns>
        public static Attachment GetNotificationCardForRejectedRequest(CompanyResponseEntity userRequestDetails, IStringLocalizer<Strings> localizer)
        {
            if (userRequestDetails == null)
            {
                throw new ArgumentNullException(nameof(userRequestDetails));
            }

            bool showRemarkField = !string.IsNullOrEmpty(userRequestDetails.ApprovalRemark);
            var formattedDateTime = userRequestDetails.ApprovedOrRejectedDate.ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.InvariantCulture);
            string dateString = string.Format(CultureInfo.InvariantCulture, localizer.GetString("DateFormat"), "{{DATE(" + formattedDateTime + ", COMPACT)}}", "{{TIME(" + formattedDateTime + ")}}");

            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 2))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveContainer
                    {
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Text = localizer.GetString("NotificationCardApprovedOrRejectedTitle"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = string.Format(CultureInfo.InvariantCulture, localizer.GetString("NotificationCardStatusText"), localizer.GetString("RejectedRequestStatusText")),
                                                Wrap = true,
                                                Color = AdaptiveTextColor.Attention,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveContainer
                    {
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Text = localizer.GetString("NotificationCardLabelText"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = userRequestDetails.QuestionLabel,
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = localizer.GetString("NotificationCardQuestion"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = userRequestDetails.QuestionText,
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = localizer.GetString("NotificationCardResponse"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = userRequestDetails.ResponseText,
                                                Wrap = true,
                                            },
                                        },
                                        Style = AdaptiveContainerStyle.Emphasis,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveContainer
                    {
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Text = localizer.GetString("NotificationCardRemark"),
                                                Wrap = true,
                                                IsVisible = showRemarkField,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = userRequestDetails.ApprovalRemark,
                                                Wrap = true,
                                                IsVisible = showRemarkField,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = string.Format(CultureInfo.InvariantCulture, localizer.GetString("RejectedCardLabelText"), dateString),
                                                Wrap = true,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                },
            };
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };

            return adaptiveCardAttachment;
        }
    }
}
