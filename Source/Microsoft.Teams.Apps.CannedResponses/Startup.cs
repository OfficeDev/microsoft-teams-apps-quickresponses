// <copyright file="Startup.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses
{
    using System;
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.AspNetCore.SpaServices.ReactDevelopmentServer;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.BotFramework;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Teams.Apps.CannedResponses.Bot;

    /// <summary>
    /// The Startup class is responsible for configuring the DI container and acts as the composition root.
    /// </summary>
    public sealed class Startup
    {
        private readonly IConfiguration configuration;

        /// <summary>
        /// Initializes a new instance of the <see cref="Startup"/> class.
        /// </summary>
        /// <param name="configuration">The environment provided configuration.</param>
        public Startup(IConfiguration configuration)
        {
            this.configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        }

        /// <summary>
        /// Configure the composition root for the application.
        /// </summary>
        /// <param name="services">The stub composition root.</param>
        /// <remarks>
        /// For more information see: https://go.microsoft.com/fwlink/?LinkID=398940.
        /// </remarks>
        #pragma warning disable CA1506 // Composition root expected to have coupling with many components.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddCredentialProviders(this.configuration);
            services.AddConfigurationSettings(this.configuration);
            services.AddHelpers(this.configuration);
            services.AddCustomJWTAuthentication(this.configuration);
            services.AddBotStates(this.configuration);

            services.AddSpaStaticFiles(configuration =>
            {
                configuration.RootPath = "ClientApp/build";
            });

            // Create the Bot Framework Adapter with error handling enabled.
            services.AddSingleton<IBotFrameworkHttpAdapter, CannedResponsesAdapterWithErrorHandler>();

            // Create the Bot Framework Adapter with error handling enabled.
            services.AddTransient<ICredentialProvider, ConfigurationCredentialProvider>();

            services.AddTransient<IBot, CannedResponsesActivityHandler>();

            // Create the Translation Middleware that will be added to the middleware pipeline in the AdapterWithErrorHandler
            services.AddSingleton<CannedResponsesActivityMiddleWare>();
            services.AddTransient(serviceProvider => (BotFrameworkAdapter)serviceProvider.GetRequiredService<IBotFrameworkHttpAdapter>());

            services.AddSingleton(new MicrosoftAppCredentials(this.configuration.GetValue<string>("MicrosoftAppId"), this.configuration.GetValue<string>("MicrosoftAppPassword")));

            services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_1);
            services.AddLocalization(this.configuration);
        }

        #pragma warning restore CA1506
        /// <summary>
        /// Configure the application request pipeline.
        /// </summary>
        /// <param name="app">The application.</param>
        /// <param name="env">Hosting Environment.</param>
        #pragma warning disable CA1822 // This method is provided by the framework
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        #pragma warning restore CA1822
        {
            app.UseRequestLocalization();
            app.UseAuthentication();
            app.UseStaticFiles();
            app.UseSpaStaticFiles();
            app.UseMvc();
            app.UseSpa(spa =>
            {
                spa.Options.SourcePath = "ClientApp";

                if (env.IsDevelopment())
                {
                    spa.UseReactDevelopmentServer(npmScript: "start");
                }
            });
        }
    }
}
