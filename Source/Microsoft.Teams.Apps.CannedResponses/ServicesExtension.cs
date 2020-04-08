// <copyright file="ServicesExtension.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using Microsoft.AspNetCore.Authentication.JwtBearer;
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Localization;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Azure;
    using Microsoft.Bot.Builder.BotFramework;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.IdentityModel.Tokens;
    using Microsoft.Teams.Apps.CannedResponses.Bot;
    using Microsoft.Teams.Apps.CannedResponses.Common.BackgroundService;
    using Microsoft.Teams.Apps.CannedResponses.Common.Interfaces;
    using Microsoft.Teams.Apps.CannedResponses.Common.Providers;
    using Microsoft.Teams.Apps.CannedResponses.Common.SearchServices;
    using Microsoft.Teams.Apps.CannedResponses.Helpers;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// Class which helps to extend ServiceCollection.
    /// </summary>
    public static class ServicesExtension
    {
        /// <summary>
        /// Adds application configuration settings to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddConfigurationSettings(this IServiceCollection services, IConfiguration configuration)
        {
            string appBaseUrl = configuration.GetValue<string>("App:AppBaseUri");

            services.Configure<CannedResponsesActivityHandlerOptions>(options =>
            {
                options.UpperCaseResponse = configuration.GetValue<bool>("UppercaseResponse");
                options.AppBaseUri = appBaseUrl;
            });

            services.Configure<BotSetting>(options =>
            {
                options.SecurityKey = configuration.GetValue<string>("App:SecurityKey");
                options.AppBaseUri = appBaseUrl;
                options.TeamIdDeepLink = configuration.GetValue<string>("Teams:TeamIdDeepLink");
                options.TenantId = configuration.GetValue<string>("App:TenantId");
            });

            services.Configure<TelemetrySetting>(options =>
            {
                options.InstrumentationKey = configuration.GetValue<string>("ApplicationInsights:InstrumentationKey");
            });

            services.Configure<StorageSetting>(options =>
            {
                options.ConnectionString = configuration.GetValue<string>("Storage:ConnectionString");
            });

            services.Configure<SearchServiceSetting>(searchServiceSettings =>
            {
                searchServiceSettings.SearchServiceName = configuration.GetValue<string>("SearchService:SearchServiceName");
                searchServiceSettings.SearchServiceQueryApiKey = configuration.GetValue<string>("SearchService:SearchServiceQueryApiKey");
                searchServiceSettings.SearchServiceAdminApiKey = configuration.GetValue<string>("SearchService:SearchServiceAdminApiKey");
                searchServiceSettings.SearchIndexingIntervalInMinutes = configuration.GetValue<string>("SearchService:SearchIndexingIntervalInMinutes");
                searchServiceSettings.ConnectionString = configuration.GetValue<string>("Storage:ConnectionString");
            });
        }

        /// <summary>
        /// Adds helpers to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddHelpers(this IServiceCollection services, IConfiguration configuration)
        {
            string appId = configuration.GetValue<string>("MicrosoftAppId");
            string appPassword = configuration.GetValue<string>("MicrosoftAppPassword");
            string appBaseUrl = configuration.GetValue<string>("App:AppBaseUri");

            ICredentialProvider credentialProvider = new SimpleCredentialProvider(
                appId: appId,
                password: appPassword);

            services.AddSingleton(credentialProvider);

            services.AddApplicationInsightsTelemetry(configuration.GetValue<string>("ApplicationInsights:InstrumentationKey"));

            services.AddSingleton<ITokenHelper, TokenHelper>();
            services.AddHostedService<UserResponseDataRefreshService>();
            services.AddSingleton<ICompanyResponseStorageProvider, CompanyResponseStorageProvider>();
            services.AddSingleton<IUserResponseStorageProvider, UserResponseStorageProvider>();
            services.AddSingleton<IUserResponseSearchService, UserResponseSearchService>();
            services.AddSingleton<ICompanyResponseSearchService, CompanyResponseSearchService>();
            services.AddSingleton<ICompanyStorageHelper, CompanyStorageHelper>();
            services.AddSingleton<IUserStorageHelper, UserStorageHelper>();
            services.AddSingleton<IMessagingExtensionHelper, MessagingExtensionHelper>();
            services.AddSingleton<IConversationStorageProvider, ConversationStorageProvider>();
        }

        /// <summary>
        /// Adds custom JWT authentication to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddCustomJWTAuthentication(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
            .AddJwtBearer(options =>
            {
                options.TokenValidationParameters = new TokenValidationParameters
                {
                    ValidateAudience = true,
                    ValidAudiences = new List<string> { configuration.GetValue<string>("App:AppBaseUri") },
                    ValidIssuers = new List<string> { configuration.GetValue<string>("App:AppBaseUri") },
                    ValidateIssuer = true,
                    ValidateIssuerSigningKey = true,
                    IssuerSigningKey = new SymmetricSecurityKey(Encoding.ASCII.GetBytes(configuration.GetValue<string>("App:SecurityKey"))),
                    RequireExpirationTime = true,
                    ValidateLifetime = true,
                    ClockSkew = TimeSpan.FromSeconds(30),
                };
            });
        }

        /// <summary>
        /// Adds user state and conversation state to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddBotStates(this IServiceCollection services, IConfiguration configuration)
        {
            // Create the User state. (Used in this bot's Dialog implementation.)
            services.AddSingleton<UserState>();

            // Create the Conversation state. (Used by the Dialog system itself.)
            services.AddSingleton<ConversationState>();

            // For conversation state.
            services.AddSingleton<IStorage>(new AzureBlobStorage(configuration.GetValue<string>("Storage:ConnectionString"), "bot-state"));
        }

        /// <summary>
        /// Adds localization.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddLocalization(this IServiceCollection services, IConfiguration configuration)
        {
            // Add i18n.
            services.AddLocalization(options => options.ResourcesPath = "Resources");

            services.Configure<RequestLocalizationOptions>(options =>
            {
                var defaultCulture = CultureInfo.GetCultureInfo(configuration.GetValue<string>("i18n:DefaultCulture"));
                var supportedCultures = configuration.GetValue<string>("i18n:SupportedCultures").Split(',')
                    .Select(culture => CultureInfo.GetCultureInfo(culture))
                    .ToList();

                options.DefaultRequestCulture = new RequestCulture(defaultCulture);
                options.SupportedCultures = supportedCultures;
                options.SupportedUICultures = supportedCultures;

                options.RequestCultureProviders = new List<IRequestCultureProvider>
                {
                    new CannedResponsesLocalizationCultureProvider(),
                };
            });
        }

        /// <summary>
        /// Adds credential providers for authentication.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddCredentialProviders(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();
            services.AddSingleton(new MicrosoftAppCredentials(configuration.GetValue<string>("MicrosoftAppId"), configuration.GetValue<string>("MicrosoftAppPassword")));
        }
    }
}
