// <copyright file="WelcomeCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;

    /// <summary>
    /// Class having method to return welcome card attachment.
    /// </summary>
    public static class WelcomeCard
    {
        /// <summary>
        /// Get welcome card attachment to show on Microsoft Teams channel scope.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL to get the logo of the app.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>User welcome card attachment.</returns>
        public static Attachment GetWelcomeCardAttachmentForTeams(string applicationBasePath, IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 2))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = new Uri($"{applicationBasePath}/Artifacts/appLogo.png"),
                                        Size = AdaptiveImageSize.Medium,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Spacing = AdaptiveSpacing.None,
                                        Text = localizer.GetString("WelcomeCardContentForAdmin"),
                                        Wrap = true,
                                        IsSubtle = true,
                                    },
                                },
                                Width = AdaptiveColumnWidth.Stretch,
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("WelcomeCardThingsContentText"),
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("WelcomeCardSuggestedContentText"),
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("WelcomeCardLabelContentText"),
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("WelcomeCardChatRequestorContentText"),
                        Wrap = true,
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
        /// Get welcome card attachment to show on Microsoft Teams personal scope.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>User welcome card attachment.</returns>
        public static Attachment GetWelcomeCardAttachmentForPersonal(string applicationBasePath, IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 2))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = new Uri($"{applicationBasePath}/Artifacts/appLogo.png"),
                                        Size = AdaptiveImageSize.Medium,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Weight = AdaptiveTextWeight.Default,
                                        Spacing = AdaptiveSpacing.None,
                                        Text = localizer.GetString("WelcomeCardTitle"),
                                        Wrap = true,
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Spacing = AdaptiveSpacing.None,
                                        Text = localizer.GetString("WelcomeCardContent"),
                                        Wrap = true,
                                        IsSubtle = true,
                                    },
                                },
                                Width = AdaptiveColumnWidth.Stretch,
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("WelcomeCardThingsContentText"),
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("WelcomeCardResponseListContentText"),
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("WelcomeCardSuggestResponseContentText"),
                        Wrap = true,
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
