// <copyright file="UserResponseSearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.SearchServices
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Azure.Search;
    using Microsoft.Azure.Search.Models;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CannedResponses.Common.Interfaces;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// User response Search service which will help in creating index, indexer and data source if it doesn't exist
    /// for indexing table which will be used for search by messaging extension.
    /// </summary>
    public class UserResponseSearchService : IUserResponseSearchService, IDisposable
    {
        /// <summary>
        /// Azure Search service index name for user response.
        /// </summary>
        private const string UserResponseIndexName = "user-response-index";

        /// <summary>
        /// Azure Search service indexer name for user response.
        /// </summary>
        private const string UserResponseIndexerName = "user-response-indexer";

        /// <summary>
        /// Azure Search service data source name for user response.
        /// </summary>
        private const string UserResponseDataSourceName = "user-response-storage";

        /// <summary>
        /// Table name where user response data will get saved.
        /// </summary>
        private const string UserResponseTableName = "UserResponseEntity";

        /// <summary>
        /// Used to initialize task.
        /// </summary>
        private readonly Lazy<Task> initializeTask;

        /// <summary>
        /// Instance of Azure Search service client.
        /// </summary>
        private readonly SearchServiceClient searchServiceClient;

        /// <summary>
        /// Instance of Azure Search index client.
        /// </summary>
        private readonly SearchIndexClient searchIndexClient;

        /// <summary>
        /// Instance of user response storage helper to update response and get information of responses.
        /// </summary>
        private readonly IUserResponseStorageProvider userResponseStorageProvider;

        /// <summary>
        /// Search indexing interval in minutes.
        /// </summary>
        private readonly int searchIndexingIntervalInMinutes;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<UserResponseSearchService> logger;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly SearchServiceSetting options;

        /// <summary>
        /// Flag: Has Dispose already been called?
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserResponseSearchService"/> class.
        /// </summary>
        /// <param name="optionsAccessor">A set of key/value application configuration properties.</param>
        /// <param name="userResponseStorageProvider">User response storage provider dependency injection.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public UserResponseSearchService(
            IOptions<SearchServiceSetting> optionsAccessor,
            IUserResponseStorageProvider userResponseStorageProvider,
            ILogger<UserResponseSearchService> logger)
        {
            optionsAccessor = optionsAccessor ?? throw new ArgumentNullException(nameof(optionsAccessor));

            this.options = optionsAccessor.Value;
            var searchServiceValue = this.options.SearchServiceName;
            this.searchServiceClient = new SearchServiceClient(
                searchServiceValue,
                new SearchCredentials(this.options.SearchServiceAdminApiKey));
            this.searchIndexClient = new SearchIndexClient(
                searchServiceValue,
                UserResponseIndexName,
                new SearchCredentials(this.options.SearchServiceQueryApiKey));
            this.searchIndexingIntervalInMinutes = Convert.ToInt32(this.options.SearchIndexingIntervalInMinutes, CultureInfo.InvariantCulture);

            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(this.options.ConnectionString));
            this.userResponseStorageProvider = userResponseStorageProvider;
            this.logger = logger;
        }

        /// <summary>
        /// Provide search result for table to be used by user based on Azure Search service.
        /// </summary>
        /// <param name="searchQuery">Keyword entered by user in messaging extension search field.</param>
        /// /// <param name="userObjectId">AAd object id of the user.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <returns>List of search results.</returns>
        public async Task<IList<UserResponseEntity>> SearchUserResponseAsync(string searchQuery, string userObjectId, int? count = null, int? skip = null)
        {
            await this.EnsureInitializedAsync();
            IList<UserResponseEntity> userResponses = new List<UserResponseEntity>();
            SearchParameters searchParameters = new SearchParameters()
            {
                // Filter by current user aad object id.
                Filter = $"UserId eq '{userObjectId}' ",
                OrderBy = new[] { "LastUpdatedDate desc" },
                Top = count ?? Constants.DefaultSearchResultCount,
                Skip = skip ?? 0,
                IncludeTotalResultCount = false,
                Select = new[] { "UserId", "ResponseId", "QuestionLabel", "QuestionText", "ResponseText", "LastUpdatedDate" },
            };

            var userResponsesResult = await this.searchIndexClient.Documents.SearchAsync<UserResponseEntity>(searchQuery, searchParameters);
            if (userResponsesResult != null)
            {
                userResponses = userResponsesResult.Results.Select(p => p.Document).ToList();
            }

            return userResponses;
        }

        /// <summary>
        /// Creates Index, Data Source and Indexer for search service.
        /// </summary>
        /// <param name="connectionString">Connection string to the data store.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task InitializeSearchServiceIndexAsync(string connectionString)
        {
            try
            {
                await this.CreateSearchIndexAsync();
                await this.CreateDataSourceAsync(connectionString);
                await this.CreateIndexerAsync();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Disponse search service instance.
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
                this.searchServiceClient.Dispose();
                this.searchIndexClient.Dispose();
            }

            this.disposed = true;
        }

        /// <summary>
        /// Create index, indexer and data source if doesn't exist.
        /// </summary>
        /// <param name="connectionString">Connection string to the data store.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task InitializeAsync(string connectionString)
        {
            try
            {
                // When there is no user response created by user and messaging extension is open, table initialization is required here before creating search index or data source or indexer.
                await this.userResponseStorageProvider.GetUserResponsesDataAsync(string.Empty);
                await this.InitializeSearchServiceIndexAsync(connectionString);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to initialize Azure Search Service: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Create index in Azure Search service if it doesn't exist.
        /// </summary>
        /// <returns><see cref="Task"/> That represents index is created if it is not created.</returns>
        private async Task CreateSearchIndexAsync()
        {
            if (await this.searchServiceClient.Indexes.ExistsAsync(UserResponseIndexName))
            {
                await this.searchServiceClient.Indexes.DeleteAsync(UserResponseIndexName);
            }

            var tableIndex = new Index()
            {
                Name = UserResponseIndexName,
                Fields = FieldBuilder.BuildForType<UserResponseEntity>(),
            };
            await this.searchServiceClient.Indexes.CreateAsync(tableIndex);
        }

        /// <summary>
        /// Create data source if it doesn't exist in Azure Search service.
        /// </summary>
        /// <param name="connectionString">Connection string to the data store.</param>
        /// <returns><see cref="Task"/> That represents data source is added to Azure Search service.</returns>
        private async Task CreateDataSourceAsync(string connectionString)
        {
            if (await this.searchServiceClient.DataSources.ExistsAsync(UserResponseDataSourceName))
            {
                return;
            }

            var dataSource = DataSource.AzureTableStorage(
                                            UserResponseDataSourceName,
                                            connectionString,
                                            UserResponseTableName);

            await this.searchServiceClient.DataSources.CreateAsync(dataSource);
        }

        /// <summary>
        /// Create indexer if it doesn't exist in Azure Search service.
        /// </summary>
        /// <returns><see cref="Task"/> That represents indexer is created if not available in Azure Search service.</returns>
        private async Task CreateIndexerAsync()
        {
            if (await this.searchServiceClient.Indexers.ExistsAsync(UserResponseIndexerName))
            {
                await this.searchServiceClient.Indexers.DeleteAsync(UserResponseIndexerName);
            }

            var indexer = new Indexer()
            {
                Name = UserResponseIndexerName,
                DataSourceName = UserResponseDataSourceName,
                TargetIndexName = UserResponseIndexName,
                Schedule = new IndexingSchedule(TimeSpan.FromMinutes(this.searchIndexingIntervalInMinutes)),
            };

            await this.searchServiceClient.Indexers.CreateAsync(indexer);
            await this.searchServiceClient.Indexers.RunAsync(UserResponseIndexerName);
        }

        /// <summary>
        /// Initialization of InitializeAsync method which will help in indexing.
        /// </summary>
        /// <returns>Represents an asynchronous operation.</returns>
        private Task EnsureInitializedAsync()
        {
            return this.initializeTask.Value;
        }
    }
}
