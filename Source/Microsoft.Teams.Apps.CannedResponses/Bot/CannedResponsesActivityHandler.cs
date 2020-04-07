// <copyright file="CannedResponsesActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Bot
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CannedResponses.Cards;
    using Microsoft.Teams.Apps.CannedResponses.Common;
    using Microsoft.Teams.Apps.CannedResponses.Common.Interfaces;
    using Microsoft.Teams.Apps.CannedResponses.Helpers;
    using Microsoft.Teams.Apps.CannedResponses.Models;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// This class is responsible for reacting to incoming events from Microsoft Teams sent from BotFramework.
    /// </summary>
    public sealed class CannedResponsesActivityHandler : TeamsActivityHandler
    {
        /// <summary>
        /// Sets the height of the task module.
        /// </summary>
        private const int TaskModuleHeight = 600;

        /// <summary>
        /// Sets the width of the task module.
        /// </summary>
        private const int TaskModuleWidth = 800;

        /// <summary>
        /// Represents the conversation type as personal.
        /// </summary>
        private const string Personal = "PERSONAL";

        /// <summary>
        ///  Represents the conversation type as channel.
        /// </summary>
        private const string Channel = "CHANNEL";

        /// <summary>
        /// Command when user adds new response in user response entity in Microsoft Azure Table storage.
        /// </summary>
        private const string AddUserResponseCommand = "AddUserResponse";

        /// <summary>
        /// Command when user suggests new response in company response entity in Microsoft Azure Table storage.
        /// </summary>
        private const string AddNewSuggestionCommand = "AddNewSuggestion";

        /// <summary>
        /// Command when user edits response in user response entity in Microsoft Azure Table storage.
        /// </summary>
        private const string EditUserResponseCommand = "EditUserResponse";

        /// <summary>
        /// Event name for user searches in your responses.
        /// </summary>
        private const string YourResponsesSearchEventName = "Your responses search";

        /// <summary>
        /// Event name for user searches in company responses.
        /// </summary>
        private const string CompanyResponsesSearchEventName = "Company responses search";

        /// <summary>
        /// Event name for approved request.
        /// </summary>
        private const string ApprovedRequestEventName = "Approved requests";

        /// <summary>
        /// Event name for rejected request.
        /// </summary>
        private const string RejectedRequestEventName = "Rejected requests";

        /// <summary>
        ///  Company responses command id when Messaging Extension invokes company response task module.
        /// </summary>
        private const string CompanyResponseTaskModuleCommandId = "CompanyResponse";

        /// <summary>
        /// Event name for user added new response.
        /// </summary>
        private const string RecordAddNewUserResponse = "Your responses - Added successfully";

        /// <summary>
        /// Event name for user edited the response.
        /// </summary>
        private const string RecordEditUserResponse = "Your responses - Edited successfully";

        /// <summary>
        /// Event name for user suggested new response.
        /// </summary>
        private const string RecordSuggestCompanyResponse = "Company responses - Suggested successfully";

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<CannedResponsesActivityHandlerOptions> options;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<CannedResponsesActivityHandler> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Instance of Application Insights Telemetry client.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Instance of user response storage provider to update response and get information of responses.
        /// </summary>
        private readonly IUserResponseStorageProvider userResponseStorageProvider;

        /// <summary>
        /// Instance of token helper for holding custom jwt token for retrieving user detail.
        /// </summary>
        private readonly ITokenHelper tokenHelper;

        /// <summary>
        /// Storage provider for working with company responses data in Microsoft Azure Table storage.
        /// </summary>
        private readonly ICompanyResponseStorageProvider companyResponseStorageProvider;

        /// <summary>
        /// Helper for working with Microsoft Azure Table Search service.
        /// </summary>
        private readonly IUserResponseSearchService userResponseSearchService;

        /// <summary>
        /// Helper for working with Microsoft Azure Table Search service.
        /// </summary>
        private readonly ICompanyResponseSearchService companyResponseSearchService;

        /// <summary>
        /// State management object for maintaining user conversation state.
        /// </summary>
        private readonly BotState userState;

        /// <summary>
        ///  Holds app credentials to send the given attachment to the specified team.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Retrieve required bot configurations.
        /// </summary>
        private readonly IOptions<BotSetting> botSetting;

        /// <summary>
        /// Application Insights Telemetry settings.
        /// </summary>
        private readonly IOptions<TelemetrySetting> telemetrySettings;

        /// <summary>
        /// Storage helper for working with user responses data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IUserStorageHelper userStorageHelper;

        /// <summary>
        /// Storage helper for working with company responses data in Microsoft Azure Table storage.
        /// </summary>
        private readonly ICompanyStorageHelper companyStorageHelper;

        /// <summary>
        /// Messaging Extension search helper for working with company and user responses data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IMessagingExtensionHelper messagingExtensionHelper;

        /// <summary>
        /// Storage helper for working with user conversation data in Microsoft Azure Table storage.
        /// </summary>
        private readonly IConversationStorageProvider conversationStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="CannedResponsesActivityHandler"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client.</param>
        /// <param name="options">>A set of key/value application configuration properties for activity handler.</param>
        /// <param name="userResponseStorageProvider">User response storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="companyResponseStorageProvider">Company response storage provider to maintain data in Microsoft Azure table storage.</param>
        /// <param name="userResponseSearchService">The user response search service dependency injection.</param>
        /// <param name="companyResponseSearchService">The company response search service dependency injection.</param>
        /// <param name="userState">State management object for maintaining user conversation state.</param>
        /// <param name="microsoftAppCredentials">Microsoft application credentials to send card to the specified team.</param>
        /// <param name="tokenHelper">Generating custom JWT token and retrieving user detail from token.</param>
        /// <param name="botSetting">>A set of key/value application configuration properties for bot.</param>
        /// <param name="telemetrySettings">>Application insights settings.</param>
        /// <param name="userStorageHelper">User storage helper dependency injection.</param>
        /// <param name="companyStorageHelper">Company storage helper dependency injection.</param>
        /// <param name="messagingExtensionHelper">Messaging Extension helper dependency injection.</param>
        /// <param name="conversationStorageProvider">Conversation storage provider to maintain data in Microsoft Azure table storage.</param>
        public CannedResponsesActivityHandler(
            ILogger<CannedResponsesActivityHandler> logger,
            IStringLocalizer<Strings> localizer,
            TelemetryClient telemetryClient,
            IOptions<CannedResponsesActivityHandlerOptions> options,
            IUserResponseStorageProvider userResponseStorageProvider,
            ICompanyResponseStorageProvider companyResponseStorageProvider,
            IUserResponseSearchService userResponseSearchService,
            ICompanyResponseSearchService companyResponseSearchService,
            UserState userState,
            MicrosoftAppCredentials microsoftAppCredentials,
            ITokenHelper tokenHelper,
            IOptions<BotSetting> botSetting,
            IOptions<TelemetrySetting> telemetrySettings,
            IUserStorageHelper userStorageHelper,
            ICompanyStorageHelper companyStorageHelper,
            IMessagingExtensionHelper messagingExtensionHelper,
            IConversationStorageProvider conversationStorageProvider)
        {
            this.logger = logger;
            this.localizer = localizer;
            this.telemetryClient = telemetryClient;
            this.options = options;
            this.userResponseStorageProvider = userResponseStorageProvider;
            this.companyResponseStorageProvider = companyResponseStorageProvider;
            this.userResponseSearchService = userResponseSearchService;
            this.companyResponseSearchService = companyResponseSearchService;
            this.userState = userState;
            this.tokenHelper = tokenHelper;
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.botSetting = botSetting;
            this.telemetrySettings = telemetrySettings;
            this.userStorageHelper = userStorageHelper;
            this.companyStorageHelper = companyStorageHelper;
            this.messagingExtensionHelper = messagingExtensionHelper;
            this.conversationStorageProvider = conversationStorageProvider;
        }

        /// <summary>
        /// Handles an incoming activity.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.activityhandler.onturnasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        public override Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnTurnAsync), turnContext);
                this.userState.SaveChangesAsync(turnContext, false, cancellationToken);
                return base.OnTurnAsync(turnContext, cancellationToken);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error at OnTurnAsync(): {ex.Message}", SeverityLevel.Error);
                this.userState.SaveChangesAsync(turnContext, false, cancellationToken);
                return base.OnTurnAsync(turnContext, cancellationToken);
            }
        }

        /// <summary>
        /// Invoked when a message activity is received from the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.activityhandler.onmessageactivityasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task OnMessageActivityAsync(
            ITurnContext<IMessageActivity> turnContext,
            CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnMessageActivityAsync), turnContext);

                var activity = turnContext.Activity;

                switch (activity.Conversation.ConversationType.ToUpperInvariant())
                {
                    case Personal:
                        await turnContext.SendActivityAsync(this.localizer.GetString("UserCustomMessage"));
                        break;

                    case Channel:
                        await this.OnMessageActivityInChannelAsync(
                            activity,
                            turnContext,
                            cancellationToken);
                        break;

                    default:
                        this.logger.LogInformation($"Received unexpected conversationType {activity.Conversation.ConversationType}", SeverityLevel.Warning);
                        break;
                }
            }
            catch (Exception ex)
            {
                await turnContext.SendActivityAsync(this.localizer.GetString("ErrorMessage"));
                this.logger.LogError(ex, $"Error processing message: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Handle message extension action fetch task received by the bot.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="action">Messaging extension action value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Response of messaging extension action.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionfetchtaskasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionFetchTaskAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnTeamsMessagingExtensionFetchTaskAsync), turnContext);

                action = action ?? throw new ArgumentNullException(nameof(action));

                var activity = turnContext.Activity;

                // Generate custom JWT token to authenticate in application API controller.
                var customAPIAuthenticationToken = this.tokenHelper.GenerateInternalAPIToken(new Uri(activity.ServiceUrl), activity.From.AadObjectId, jwtExpiryMinutes: 60);

                return this.GetTaskModuleBasedOnCommand(action, customAPIAuthenticationToken);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while fetching task module received by the bot which is invoked through ME.");
                throw;
            }
        }

        /// <summary>
        /// Invoked when the user submits a response/suggests a response/updates a response.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="action">Messaging extension action commands.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionsubmitactionasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionAction action,
            CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnTeamsMessagingExtensionSubmitActionAsync), turnContext);

                action = action ?? throw new ArgumentNullException(nameof(action));

                var activity = turnContext.Activity;
                var userRequestDetails = JsonConvert.DeserializeObject<AddUserResponseRequestDetail>(action.Data.ToString());

                // Generate custom JWT token to authenticate in application API controller.
                var customAPIAuthenticationToken = this.tokenHelper.GenerateInternalAPIToken(new Uri(activity.ServiceUrl), activity.From.AadObjectId, jwtExpiryMinutes: 60);

                switch (userRequestDetails.CommandContext)
                {
                    case AddUserResponseCommand:
                        return await this.AddUserResponseResultAsync(turnContext, userRequestDetails, customAPIAuthenticationToken);

                    case AddNewSuggestionCommand:
                        return await this.AddNewSuggestionResultAsync(turnContext, userRequestDetails, customAPIAuthenticationToken, cancellationToken);

                    case EditUserResponseCommand:
                        return await this.EditUserResponseResultAsync(turnContext, userRequestDetails, customAPIAuthenticationToken);
                }

                return null;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error at OnTeamsMessagingExtensionSubmitActionAsync(): {ex.Message}", SeverityLevel.Error);
                await turnContext.SendActivityAsync(this.localizer.GetString("ErrorMessage"));
                throw;
            }
        }

        /// <summary>
        /// Invoked when Bot/Messaging Extension is installed in team to send welcome card.
        /// </summary>
        /// <param name="membersAdded">A list of all the members added to the conversation, as described by the conversation update activity.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents welcome card when bot is added first time by user.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.activityhandler.onmembersaddedasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnMembersAddedAsync), turnContext);

                var activity = turnContext.Activity;
                this.logger.LogInformation($"conversationType: {activity.Conversation.ConversationType}, membersAdded: {activity.MembersAdded?.Count}, membersRemoved: {activity.MembersRemoved?.Count}");

                if (activity.Conversation.ConversationType.Equals(Personal, StringComparison.OrdinalIgnoreCase))
                {
                    if (activity.MembersAdded.FirstOrDefault(member => member.Id != activity.Recipient.Id) != null)
                    {
                        this.logger.LogInformation($"Bot added {activity.Conversation.Id}");
                        var userStateAccessors = this.userState.CreateProperty<UserConversationState>(nameof(UserConversationState));
                        var userConversationState = await userStateAccessors.GetAsync(turnContext, () => new UserConversationState());
                        if (userConversationState?.IsWelcomeCardSent == null || userConversationState?.IsWelcomeCardSent == false)
                        {
                            userConversationState.IsWelcomeCardSent = true;
                            await userStateAccessors.SetAsync(turnContext, userConversationState);
                            var userWelcomeCardAttachment = WelcomeCard.GetWelcomeCardAttachmentForPersonal(this.options.Value.AppBaseUri, localizer: this.localizer);
                            await turnContext.SendActivityAsync(MessageFactory.Attachment(userWelcomeCardAttachment));

                            // Store user conversation id in table storage for future use.
                            ConversationEntity conversationEntity = new ConversationEntity { ConversationId = activity.Conversation.Id, UserId = activity.From.AadObjectId };
                            bool operationStatus = await this.conversationStorageProvider.AddConversationEntityAsync(conversationEntity);
                            if (!operationStatus)
                            {
                                this.logger.LogInformation($"Unable to add conversation data in table storage.");
                            }
                        }
                    }
                    else
                    {
                        this.logger.LogError("User data could not be found at OnMembersAddedAsync().");
                    }
                }
                else
                {
                    if (activity.MembersAdded.FirstOrDefault(member => member.Id == activity.Recipient.Id) != null)
                    {
                        this.logger.LogInformation($"Bot added {activity.Conversation.Id}");
                        var userWelcomeCardAttachment = WelcomeCard.GetWelcomeCardAttachmentForTeams(this.options.Value.AppBaseUri, localizer: this.localizer);
                        await turnContext.SendActivityAsync(MessageFactory.Attachment(userWelcomeCardAttachment));
                    }
                    else
                    {
                        this.logger.LogError("User data could not be found at OnMembersAddedAsync().");
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError("Exception occurred while sending the welcome card", ex);
                throw;
            }
        }

        /// <summary>
        /// Invoked when the user opens the messaging extension or searching any content in it.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="query">Contains messaging extension query keywords.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Messaging extension response object to fill compose extension section.</returns>
        /// <remarks>
        /// https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionqueryasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionQuery query,
            CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            this.RecordEvent(nameof(this.OnTeamsMessagingExtensionQueryAsync), turnContext);

            var activity = turnContext.Activity;
            try
            {
                var messagingExtensionQuery = JsonConvert.DeserializeObject<MessagingExtensionQuery>(activity.Value.ToString());
                var searchQuery = this.messagingExtensionHelper.GetSearchQueryString(messagingExtensionQuery);

                if (!string.IsNullOrEmpty(searchQuery))
                {
                    switch (messagingExtensionQuery.CommandId)
                    {
                        case Constants.YourResponseCommandId:
                            // Tracking the user search keywords from the messaging extension.
                            this.RecordEvent(YourResponsesSearchEventName, turnContext);
                            break;

                        case Constants.CompanyResponseCommandId:
                            // Tracking the company search keywords from the messaging extension.
                            this.RecordEvent(CompanyResponsesSearchEventName, turnContext);
                            break;
                    }
                }

                return new MessagingExtensionResponse
                {
                    ComposeExtension = await this.messagingExtensionHelper.GetSearchResultAsync(searchQuery, messagingExtensionQuery.CommandId, messagingExtensionQuery.QueryOptions.Count, messagingExtensionQuery.QueryOptions.Skip, activity.From.AadObjectId, localizer: this.localizer),
                };
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to handle the messaging extension command {activity.Name}: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        private MessagingExtensionActionResponse GetTaskModuleBasedOnCommand(MessagingExtensionAction action, string customAPIAuthenticationToken)
        {
            if (action.CommandId == CompanyResponseTaskModuleCommandId)
            {
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Url = $"{this.options.Value.AppBaseUri}/company-responses?token={customAPIAuthenticationToken}&telemetry={this.telemetrySettings.Value.InstrumentationKey}&theme=" + "{theme}&locale=" + "{locale}",
                            Height = TaskModuleHeight,
                            Width = TaskModuleWidth,
                            Title = this.localizer.GetString("ManageYourResponsesTitleText"),
                        },
                    },
                };
            }
            else
            {
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Url = $"{this.options.Value.AppBaseUri}/user-responses?token={customAPIAuthenticationToken}&telemetry={this.telemetrySettings.Value.InstrumentationKey}&theme=" + "{theme}&locale=" + "{locale}",
                            Height = TaskModuleHeight,
                            Width = TaskModuleWidth,
                            Title = this.localizer.GetString("ManageYourResponsesTitleText"),
                        },
                    },
                };
            }
        }

        /// <summary>
        /// Add a new response in user response entity in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="userRequestDetails">User response new request details object used to send new request data.</param>
        /// <param name="customAPIAuthenticationToken">Generate JWT token used by client application to authenticate HTTP calls with API.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task<MessagingExtensionActionResponse> AddUserResponseResultAsync(
            ITurnContext<IInvokeActivity> turnContext,
            AddUserResponseRequestDetail userRequestDetails,
            string customAPIAuthenticationToken)
        {
            var isAddNewSuccess = await this.userStorageHelper.AddNewUserRequestDetailsAsync(turnContext.Activity.From.AadObjectId, userRequestDetails);

            if (isAddNewSuccess)
            {
                // Tracking for your response added request.
                this.RecordEvent(RecordAddNewUserResponse, turnContext);
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Url = $"{this.options.Value.AppBaseUri}/response-message?token={customAPIAuthenticationToken}&status=addSuccess&message={this.localizer.GetString("AddUserResponseSuccessMessage")}&telemetry=${this.telemetrySettings.Value.InstrumentationKey}&theme=" + "{theme}&locale=" + "{locale}",
                            Height = TaskModuleHeight,
                            Width = TaskModuleWidth,
                            Title = this.localizer.GetString("ManageYourResponsesTitleText"),
                        },
                    },
                };
            }
            else
            {
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Url = $"{this.options.Value.AppBaseUri}/response-message?token={customAPIAuthenticationToken}&status=addFailed&message={this.localizer.GetString("AddUserResponseFailedMessage")}&theme=" + "{theme}&locale=" + "{locale}",
                            Height = TaskModuleHeight,
                            Width = TaskModuleWidth,
                            Title = this.localizer.GetString("ManageYourResponsesTitleText"),
                        },
                    },
                };
            }
        }

        /// <summary>
        /// Add a new suggestion in company response entity in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="userRequestDetails">User response new request details object used to send new request data.</param>
        /// <param name="customAPIAuthenticationToken">Generate JWT token used by client application to authenticate HTTP calls with API.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task<MessagingExtensionActionResponse> AddNewSuggestionResultAsync(
            ITurnContext<IInvokeActivity> turnContext,
            AddUserResponseRequestDetail userRequestDetails,
            string customAPIAuthenticationToken,
            CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity;
            var companyResponseEntity = await this.companyStorageHelper.AddNewSuggestionAsync(activity, userRequestDetails);

            // Parse team channel deep link URL and get team id.
            var teamId = AdaptiveCardHelper.ParseTeamIdFromDeepLink(this.botSetting.Value.TeamIdDeepLink);

            if (string.IsNullOrEmpty(teamId))
            {
                throw new NullReferenceException("Provided team details seems to incorrect, please reach out to the Admin.");
            }

            var isAddSuggestionSuccess = await this.companyResponseStorageProvider.UpsertConverationStateAsync(companyResponseEntity);

            if (isAddSuggestionSuccess)
            {
                // Tracking for company response suggested request.
                this.RecordEvent(RecordSuggestCompanyResponse, turnContext);
                var attachment = AdminCard.GetNewResponseRequestCard(companyResponseEntity, localizer: this.localizer);
                var resourceResponse = await this.SendCardToTeamAsync(turnContext, attachment, teamId, cancellationToken);
                companyResponseEntity.ActivityId = resourceResponse.ActivityId;
                await this.companyResponseStorageProvider.UpsertConverationStateAsync(companyResponseEntity);

                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Url = $"{this.options.Value.AppBaseUri}/response-message?token={customAPIAuthenticationToken}&status=addSuccess&isCompanyResponse=true&message={this.localizer.GetString("AddNewSuggestionSuccessMessage")}&telemetry=${this.telemetrySettings.Value.InstrumentationKey}&theme=" + "{theme}&locale=" + "{locale}",
                            Height = TaskModuleHeight,
                            Width = TaskModuleWidth,
                            Title = this.localizer.GetString("ManageYourResponsesTitleText"),
                        },
                    },
                };
            }
            else
            {
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Url = $"{this.options.Value.AppBaseUri}/response-message?token={customAPIAuthenticationToken}&status=editFailed&isCompanyResponse=true&message={this.localizer.GetString("AddNewSuggestionFailedMessage")}&theme=" + "{theme}&locale=" + "{locale}",
                            Height = TaskModuleHeight,
                            Width = TaskModuleWidth,
                            Title = this.localizer.GetString("ManageYourResponsesTitleText"),
                        },
                    },
                };
            }
        }

        /// <summary>
        /// Edit a response in user response entity in Microsoft Azure Table storage.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="userRequestDetails">User response new request details object used to send new request data.</param>
        /// <param name="customAPIAuthenticationToken">Generate JWT token used by client app to authenticate HTTP calls with API.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task<MessagingExtensionActionResponse> EditUserResponseResultAsync(
            ITurnContext<IInvokeActivity> turnContext,
            AddUserResponseRequestDetail userRequestDetails,
            string customAPIAuthenticationToken)
        {
            var isEditRequestSuccess = await this.userStorageHelper.UpdateUserRequestDetailsAsync(turnContext.Activity, userRequestDetails);
            if (isEditRequestSuccess)
            {
                // Tracking for your response edit.
                this.RecordEvent(RecordEditUserResponse, turnContext);
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Url = $"{this.options.Value.AppBaseUri}/response-message?token={customAPIAuthenticationToken}&status=editSuccess&message={this.localizer.GetString("EditUserResponseSuccessMessage")}&telemetry=${this.telemetrySettings.Value.InstrumentationKey}&theme=" + "{theme}&locale=" + "{locale}",
                            Height = TaskModuleHeight,
                            Width = TaskModuleWidth,
                            Title = this.localizer.GetString("ManageYourResponsesTitleText"),
                        },
                    },
                };
            }
            else
            {
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Url = $"{this.options.Value.AppBaseUri}/response-message?token={customAPIAuthenticationToken}&status=editFailed&message={this.localizer.GetString("EditUserResponseFailedMessage")}&theme=" + "{theme}&locale=" + "{locale}",
                            Height = TaskModuleHeight,
                            Width = TaskModuleWidth,
                            Title = this.localizer.GetString("ManageYourResponsesTitleText"),
                        },
                    },
                };
            }
        }

        /// <summary>
        /// Records event data to Application Insights telemetry client
        /// </summary>
        /// <param name="eventName">Name of the event.</param>
        /// <param name="turnContext">Provides context for a turn in a bot.</param>
        private void RecordEvent(string eventName, ITurnContext turnContext)
        {
            this.telemetryClient.TrackEvent(eventName, new Dictionary<string, string>
            {
                { "userId", turnContext.Activity.From.AadObjectId },
                { "tenantId", turnContext.Activity.Conversation.TenantId },
            });
        }

        /// <summary>
        /// Send the given attachment to the specified team.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cardToSend">The card to send.</param>
        /// <param name="teamId">Team id to which the message is being sent.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns><see cref="Task"/>That resolves to a <see cref="ConversationResourceResponse"/>Send a attachment.</returns>
        private async Task<ConversationResourceResponse> SendCardToTeamAsync(
            ITurnContext turnContext,
            Attachment cardToSend,
            string teamId,
            CancellationToken cancellationToken)
        {
            try
            {
                var conversationParameters = new ConversationParameters
                {
                    Activity = (Activity)MessageFactory.Attachment(cardToSend),
                    ChannelData = new TeamsChannelData { Channel = new ChannelInfo(teamId) },
                };

                var taskCompletionSource = new TaskCompletionSource<ConversationResourceResponse>();
                await ((BotFrameworkAdapter)turnContext.Adapter).CreateConversationAsync(
                    null,
                    turnContext.Activity.ServiceUrl,
                    this.microsoftAppCredentials,
                    conversationParameters,
                    (newTurnContext, newCancellationToken) =>
                    {
                        var activity = newTurnContext.Activity;
                        taskCompletionSource.SetResult(new ConversationResourceResponse
                        {
                            Id = activity.Conversation.Id,
                            ActivityId = activity.Id,
                            ServiceUrl = activity.ServiceUrl,
                        });
                        return Task.CompletedTask;
                    },
                    cancellationToken);

                return await taskCompletionSource.Task;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while sending card to Admin team channel: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Handle message activity in channel.
        /// </summary>
        /// <param name="message">A message in a conversation.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task OnMessageActivityInChannelAsync(
            IMessageActivity message,
            ITurnContext<IMessageActivity> turnContext,
            CancellationToken cancellationToken)
        {
            try
            {
                if (message.Value == null)
                {
                    await turnContext.SendActivityAsync(this.localizer.GetString("ErrorWhenMessageInChannel"));
                    return;
                }

                IMessageActivity userNotification = null;
                CompanyResponseEntity companyResponseEntity = null;
                var cardPostedData = ((JObject)message.Value).ToObject<AdaptiveSubmitActionData>();
                var text = cardPostedData.AdaptiveCardActions.Text;
                var activity = turnContext.Activity;

                switch (text)
                {
                    case Constants.ApproveCommand:

                        if (string.IsNullOrEmpty(cardPostedData.UpdatedQuestionCategory) || string.IsNullOrEmpty(cardPostedData.UpdatedQuestionText) || string.IsNullOrEmpty(cardPostedData.UpdatedResponseText))
                        {
                            companyResponseEntity = this.companyResponseStorageProvider.GetCompanyResponseEntityAsync(cardPostedData.ResponseId).GetAwaiter().GetResult();
                            var attachment = AdminCard.GetNewResponseRequestCard(companyResponseEntity, localizer: this.localizer, emptyApproveField: true);
                            await AdaptiveCardHelper.RefreshCardAsync(turnContext, companyResponseEntity.ActivityId, attachment);
                            return;
                        }

                        companyResponseEntity = this.companyStorageHelper.AddApprovedData(cardPostedData, activity.From.Name, activity.From.AadObjectId);
                        var approveRequestResult = this.companyResponseStorageProvider.UpsertConverationStateAsync(companyResponseEntity).GetAwaiter().GetResult();

                        if (approveRequestResult)
                        {
                            // Refresh the approved card in channel.
                            var attachment = AdminCard.GetRefreshedCardForApprovedRequest(companyResponseEntity, activity.From.Name, localizer: this.localizer);
                            await AdaptiveCardHelper.RefreshCardAsync(turnContext, companyResponseEntity.ActivityId, attachment);

                            // Get user notification attachment and send it to user for approved request.
                            userNotification = MessageFactory.Attachment(UserCard.GetNotificationCardForApprovedRequest(companyResponseEntity, localizer: this.localizer));

                            var result = await this.conversationStorageProvider.GetConversationEntityAsync(companyResponseEntity.UserId);
                            if (result != null)
                            {
                                await AdaptiveCardHelper.SendNotificationCardAsync(turnContext, userNotification, result.ConversationId, cancellationToken);

                                // Tracking for number of requests approved.
                                this.RecordEvent(ApprovedRequestEventName, turnContext);
                            }
                            else
                            {
                                this.logger.LogInformation("Unable to send notification card for approved request because conversation id is null.");
                            }
                        }
                        else
                        {
                            this.logger.LogInformation("Unable to approve the request.");
                        }

                        break;

                    case Constants.RejectCommand:

                        companyResponseEntity = this.companyStorageHelper.AddRejectedData(cardPostedData, activity.From.Name, activity.From.AadObjectId);
                        var rejectRequestResult = this.companyResponseStorageProvider.UpsertConverationStateAsync(companyResponseEntity).GetAwaiter().GetResult();

                        if (rejectRequestResult)
                        {
                            // Get user notification rejected card attachment.
                            var attachment = AdminCard.GetRefreshedCardForRejectedRequest(companyResponseEntity, activity.From.Name, localizer: this.localizer);
                            await AdaptiveCardHelper.RefreshCardAsync(turnContext, companyResponseEntity.ActivityId, attachment);

                            // Send end user notification for approved request.
                            userNotification = MessageFactory.Attachment(UserCard.GetNotificationCardForRejectedRequest(companyResponseEntity, localizer: this.localizer));

                            var result = await this.conversationStorageProvider.GetConversationEntityAsync(companyResponseEntity.UserId);
                            if (result != null)
                            {
                                await AdaptiveCardHelper.SendNotificationCardAsync(turnContext, userNotification, result.ConversationId, cancellationToken);

                                // Tracking for number of requests rejected.
                                this.RecordEvent(RejectedRequestEventName, turnContext);
                            }
                            else
                            {
                                this.logger.LogInformation("Unable to send notification card for rejected request because conversation id is null.");
                            }
                        }
                        else
                        {
                            this.logger.LogInformation("Unable to reject the request.");
                        }

                        return;

                    default:
                        this.logger.LogInformation("Unrecognized input in channel");
                        break;
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error processing message: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }
    }
}