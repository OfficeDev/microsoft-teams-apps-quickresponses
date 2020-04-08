// <copyright file="UserResponseDataRefreshService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.BackgroundService
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CannedResponses.Common.Interfaces;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// This class inherits IHostedService and implements the methods related to background tasks to re-create Azure Search service related resources like: Indexes and Indexer if timer matched.
    /// </summary>
    public class UserResponseDataRefreshService : IHostedService, IDisposable
    {
        private readonly ILogger<UserResponseDataRefreshService> logger;

        /// <summary>
        /// Helper for working with Microsoft Azure Table storage.
        /// </summary>
        private readonly IUserResponseSearchService userResponseSearchService;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly SearchServiceSetting options;

        private System.Timers.Timer timer;
        private int executionCount = 0;

        // Flag: Has Dispose already been called?
        private bool disposed = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserResponseDataRefreshService"/> class.
        /// BackgroundService class that inherits IHostedService and implements the methods related to re-create Azure Search service related resources like: Indexes and Indexer tasks.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="userResponseSearchService">The user response search service dependency injection.</param>
        /// <param name="optionsAccessor">A set of key/value application configuration properties.</param>
        public UserResponseDataRefreshService(
            ILogger<UserResponseDataRefreshService> logger,
            IUserResponseSearchService userResponseSearchService,
            IOptions<SearchServiceSetting> optionsAccessor)
        {
            this.logger = logger;
            this.userResponseSearchService = userResponseSearchService;
            this.options = optionsAccessor?.Value;
        }

        /// <summary>
        /// Method to start the background task when application starts.
        /// </summary>
        /// <param name="cancellationToken">Signals cancellation to the executing method.</param>
        /// <returns>A task instance.</returns>
        public async Task StartAsync(CancellationToken cancellationToken)
        {
            try
            {
                this.logger.LogInformation("Search service indexes, indexer re-creation Hosted Service is running.");
                await this.ScheduleAzureSearchResourcesCreationAsync();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while running the background service to refresh the data for user responses): {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Triggered when the host is performing a graceful shutdown.
        /// </summary>
        /// <param name="cancellationToken">Signals cancellation to the executing method.</param>
        /// <returns>A task instance.</returns>
        public async Task StopAsync(CancellationToken cancellationToken)
        {
            this.logger.LogInformation("Search service indexes, indexer re-creation Hosted Service is stopping.");
            await Task.CompletedTask;
        }

        /// <summary>
        /// This code added to correctly implement the disposable pattern.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Protected implementation of Dispose pattern.
        /// </summary>
        /// <param name="disposing">True if already disposed else false.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (this.disposed)
            {
                return;
            }

            if (disposing)
            {
                this.timer.Dispose();
            }

            this.disposed = true;
        }

        /// <summary>
        /// Set the timer and enqueue to re-create Azure Search service related resources like: Indexes and Indexer if timer matched.
        /// </summary>
        /// <returns>A task that Enqueue re-create Azure Search service resources task.</returns>
        private async Task ScheduleAzureSearchResourcesCreationAsync()
        {
            var count = Interlocked.Increment(ref this.executionCount);
            this.logger.LogInformation("Search service indexes, indexer re-creation Hosted Service is working. Count: {Count}", count);

            // Run after every 10 minute(s).
            this.timer = new System.Timers.Timer(600000);

            this.timer.Elapsed += async (sender, args) =>
            {
                this.logger.LogInformation($"Timer matched to re-create Search service indexes, indexer at timer value : {this.timer}");
                this.timer.Stop();  // reset timer
                await this.RecreateAzureSearchResourcesAsync(); // Queue the re-create Search service indexes, indexer task.
                await this.ScheduleAzureSearchResourcesCreationAsync();    // reschedule next
            };

            this.timer.Start();
        }

        /// <summary>
        /// Method invokes task to re-create the Search service indexes, indexer.
        /// </summary>
        /// <returns>A task that create Search service indexes, indexer.</returns>
        private async Task RecreateAzureSearchResourcesAsync()
        {
            this.logger.LogInformation("Search service indexes, indexer re-creation task queued.");
            await this.userResponseSearchService.InitializeSearchServiceIndexAsync(this.options.ConnectionString); // re-create the Search service indexes, indexer.
        }
    }
}
