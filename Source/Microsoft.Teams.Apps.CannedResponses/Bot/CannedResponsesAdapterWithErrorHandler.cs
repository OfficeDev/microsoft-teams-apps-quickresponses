// <copyright file="CannedResponsesAdapterWithErrorHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Bot
{
    using System;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.CannedResponses;

    /// <summary>
    /// Implements Error Handler.
    /// </summary>
    public class CannedResponsesAdapterWithErrorHandler : BotFrameworkHttpAdapter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CannedResponsesAdapterWithErrorHandler"/> class.
        /// </summary>
        /// <param name="configuration">Application configurations.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="cannedResponsesActivityMiddleWare">Represents middle ware that can operate on incoming activities.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="conversationState">conversationState.</param>
        public CannedResponsesAdapterWithErrorHandler(
            IConfiguration configuration,
            ILogger<IBotFrameworkHttpAdapter> logger,
            CannedResponsesActivityMiddleWare cannedResponsesActivityMiddleWare,
            IStringLocalizer<Strings> localizer,
            ConversationState conversationState = null)
            : base(configuration)
        {
            if (cannedResponsesActivityMiddleWare == null)
            {
                throw new NullReferenceException(nameof(cannedResponsesActivityMiddleWare));
            }

            // Add activity middleware to the adapter's middleware pipeline
            this.Use(cannedResponsesActivityMiddleWare);

            this.OnTurnError = async (turnContext, exception) =>
            {
                // Log any leaked exception from the application.
                logger.LogError(exception, $"Exception caught : {exception.Message}");

                // Send a catch-all apology to the user.
                await turnContext.SendActivityAsync(localizer.GetString("ErrorMessage"));

                if (conversationState != null)
                {
                    try
                    {
                        // Delete the conversationState for the current conversation to prevent the
                        // bot from getting stuck in a error-loop caused by being in a bad state.
                        // ConversationState should be thought of as similar to "cookie-state" in a Web pages.
                        await conversationState.DeleteAsync(turnContext);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, $"Exception caught on attempting to Delete ConversationState : {ex.Message}");
                    }
                }
            };
        }
    }
}
