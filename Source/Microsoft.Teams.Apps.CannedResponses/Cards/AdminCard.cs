// <copyright file="AdminCard.cs" company="Microsoft">
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
    /// Class having method to return Admin related card attachment.
    /// </summary>
    public static class AdminCard
    {
        /// <summary>
        /// Maximum limit for category field on approve toggle card.
        /// </summary>
        private const int ApproveToggleCardCategoryFieldMaxLimit = 200;

        /// <summary>
        /// Maximum limit for question field on approve toggle card.
        /// </summary>
        private const int ApproveToggleCardQuestionFieldMaxLimit = 500;

        /// <summary>
        /// Maximum limit for response field on approve toggle card.
        /// </summary>
        private const int ApproveToggleCardResponseFieldMaxLimit = 500;

        /// <summary>
        /// Maximum limit for category field on approve toggle card.
        /// </summary>
        private const int RejectToggleCardCommentFieldMaxLimit = 200;

        /// <summary>
        /// Sets approval status as approved whenever new suggestion is submitted.
        /// </summary>
        private const string ApprovedRequestStatus = "Approved";

        /// <summary>
        /// Sets approval status as rejected whenever new suggestion is submitted.
        /// </summary>
        private const string RejectedRequestStatus = "Rejected";

        /// <summary>
        /// Get new response request card to create new response.
        /// </summary>
        /// <param name="userRequestDetails">User request details object.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="emptyApproveField">Show error message if SME has missed to fill any field while approving the response.</param>
        /// <returns>An new response request card attachment.</returns>
        public static Attachment GetNewResponseRequestCard(CompanyResponseEntity userRequestDetails, IStringLocalizer<Strings> localizer, bool emptyApproveField = false)
        {
            if (userRequestDetails == null)
            {
                throw new ArgumentNullException(nameof(userRequestDetails));
            }

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
                                                Text = localizer.GetString("NotificationRequestCardContentText"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = string.Format(CultureInfo.InvariantCulture, localizer.GetString("NotificationCardRequestText"), userRequestDetails.CreatedBy),
                                                Wrap = true,
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
                            new AdaptiveTextBlock
                            {
                                Text = localizer.GetString("ErrorMessageOnApprove"),
                                Wrap = true,
                                IsVisible = emptyApproveField,
                                Color = AdaptiveTextColor.Attention,
                            },
                        },
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveShowCardAction
                    {
                        Title = localizer.GetString("ApproveButtonTitle"),
                        Card = GetApproveCard(userRequestDetails, localizer: localizer),
                    },
                    new AdaptiveShowCardAction
                    {
                        Title = localizer.GetString("RejectButtonTitle"),
                        Card = GetRejectCard(userRequestDetails, localizer: localizer),
                    },
                    new AdaptiveOpenUrlAction
                    {
                        Title = string.Format(CultureInfo.InvariantCulture, localizer.GetString("ChatTextButton"), userRequestDetails.CreatedBy),
                        UrlString = $"https://teams.microsoft.com/l/chat/0/0?users={Uri.EscapeDataString(userRequestDetails.UserPrincipalName)}",
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
        /// Get refreshed card for approved request.
        /// </summary>
        /// <param name="userRequestDetails">User request details object.</param>
        /// <param name="approvedBy">User name who approved the request.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>An attachment.</returns>
        public static Attachment GetRefreshedCardForApprovedRequest(CompanyResponseEntity userRequestDetails, string approvedBy, IStringLocalizer<Strings> localizer)
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
                                                Text = localizer.GetString("ApprovedRequestRefreshCardTitle"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = string.Format(CultureInfo.InvariantCulture, localizer.GetString("RefreshedNotificationCardRequestText"), userRequestDetails.CreatedBy),
                                                Wrap = true,
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
                                                Text = string.Format(CultureInfo.InvariantCulture, localizer.GetString("ApprovedAdminCardLabelText"), dateString, approvedBy),
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
        ///  Get refreshed card for rejected request.
        /// </summary>
        /// <param name="userRequestDetails">User request details object.</param>
        /// <param name="rejectedBy">User name who approved the request.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>An attachment.</returns>
        public static Attachment GetRefreshedCardForRejectedRequest(CompanyResponseEntity userRequestDetails, string rejectedBy, IStringLocalizer<Strings> localizer)
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
                                                Text = localizer.GetString("NotificationCardRequestReject"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = string.Format(CultureInfo.InvariantCulture, localizer.GetString("RefreshedNotificationCardRequestText"), userRequestDetails.CreatedBy),
                                                Wrap = true,
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
                                                Text = string.Format(CultureInfo.InvariantCulture, localizer.GetString("RejectedAdminCardLabelText"), dateString, rejectedBy),
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
        /// Show approve adaptive card on new response request card.
        /// </summary>
        /// <param name="userRequestDetails">User request details object.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>An approve card to show on new response request card.</returns>
        private static AdaptiveCard GetApproveCard(CompanyResponseEntity userRequestDetails, IStringLocalizer<Strings> localizer)
        {
            return new AdaptiveCard(new AdaptiveSchemaVersion(1, 2))
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
                                                Text = localizer.GetString("ApproveToggleCardCategoryTitle"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextInput
                                            {
                                               Id = "updatedquestioncategory",
                                               Value = userRequestDetails.QuestionLabel,
                                               Placeholder = localizer.GetString("ApproveToggleCardLabelPlaceholder"),
                                               MaxLength = ApproveToggleCardCategoryFieldMaxLimit,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = localizer.GetString("ApproveToggleCardQuestionsTitle"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextInput
                                            {
                                               Id = "updatedquestiontext",
                                               Value = userRequestDetails.QuestionText,
                                               Placeholder = localizer.GetString("ApproveToggleCardQuestionsPlaceholder"),
                                               MaxLength = ApproveToggleCardQuestionFieldMaxLimit,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = localizer.GetString("ApproveToggleCardQuestionsPlaceholder"),
                                                Wrap = true,
                                                HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = localizer.GetString("ApproveToggleCardResponseTitle"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextInput
                                            {
                                               Id = "updatedresponsetext",
                                               Value = userRequestDetails.ResponseText,
                                               Placeholder = localizer.GetString("ApproveToggleCardResponsePlaceholder"),
                                               MaxLength = ApproveToggleCardResponseFieldMaxLimit,
                                               IsMultiline = true,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("ApproveToggleCardSubmitButtonTitle"),
                        Data = new AdaptiveSubmitActionData
                        {
                            AdaptiveCardActions = new CardAction
                            {
                                Type = ActionTypes.MessageBack,
                                Text = Constants.ApproveCommand,
                            },
                            ResponseId = userRequestDetails.ResponseId,
                            ApprovalStatus = ApprovedRequestStatus,
                        },
                    },
                },
            };
        }

        /// <summary>
        /// Show reject adaptive card on new response request card.
        /// </summary>
        /// <param name="userRequestDetails">User request details object.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>An reject card to show on new response request card.</returns>
        private static AdaptiveCard GetRejectCard(CompanyResponseEntity userRequestDetails, IStringLocalizer<Strings> localizer)
        {
            return new AdaptiveCard(new AdaptiveSchemaVersion(1, 2))
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
                                                Text = localizer.GetString("RejectToggleCardRemarkTitle"),
                                                Wrap = true,
                                            },
                                            new AdaptiveTextInput
                                            {
                                                Id = "approvalremark",
                                                Placeholder = localizer.GetString("RejectToggleCardRemarkPlaceholder"),
                                                IsMultiline = true,
                                                MaxLength = RejectToggleCardCommentFieldMaxLimit,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("RejectToggleCardSubmitButtonTitle"),
                        Data = new AdaptiveSubmitActionData
                        {
                            AdaptiveCardActions = new CardAction
                            {
                                Type = ActionTypes.MessageBack,
                                Text = Constants.RejectCommand,
                            },
                            ResponseId = userRequestDetails.ResponseId,
                            ApprovalStatus = RejectedRequestStatus,
                        },
                    },
                },
            };
        }
    }
}
