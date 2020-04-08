// <copyright file="CompanyResponseSearchService.cs" company="Microsoft">
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
    /// Company response Search service class which will help in creating index, indexer and data source if it doesn't exist
    /// for indexing table which will be used for search by messaging extension.
    /// </summary>
    public class CompanyResponseSearchService : ICompanyResponseSearchService, IDisposable
    {
        /// <summary>
        /// Azure Search service index name for company response entity.
        /// </summary>
        private const string CompanyResponseIndexName = "company-response-index";

        /// <summary>
        /// Azure Search service indexer name for company response entity.
        /// </summary>
        private const string CompanyResponseIndexerName = "company-response-indexer";

        /// <summary>
        /// Azure Search service data source name for company response entity.
        /// </summary>
        private const string CompanyResponseDataSourceName = "company-response-storage";

        /// <summary>
        /// Table name where company response data will get saved.
        /// </summary>
        private const string CompanyResponseTableName = "CompanyResponseEntity";

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
        /// Storage provider for working with company responses data in Microsoft Azure Table storage.
        /// </summary>
        private readonly ICompanyResponseStorageProvider companyResponseStorageProvider;

        /// <summary>
        /// Search indexing interval in minutes.
        /// </summary>
        private readonly int searchIndexingIntervalInMinutes;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<CompanyResponseSearchService> logger;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly SearchServiceSetting options;

        /// <summary>
        /// Flag: Has Dispose already been called?
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="CompanyResponseSearchService"/> class.
        /// </summary>
        /// <param name="optionsAccessor">A set of key/value application configuration properties.</param>
        /// <param name="companyResponseStorageProvider">Company response storage provider dependency injection.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public CompanyResponseSearchService(
            IOptions<SearchServiceSetting> optionsAccessor,
            ICompanyResponseStorageProvider companyResponseStorageProvider,
            ILogger<CompanyResponseSearchService> logger)
        {
            optionsAccessor = optionsAccessor ?? throw new ArgumentNullException(nameof(optionsAccessor));

            this.options = optionsAccessor.Value;
            var searchServiceValue = this.options.SearchServiceName;
            this.searchServiceClient = new SearchServiceClient(
                searchServiceValue,
                new SearchCredentials(this.options.SearchServiceAdminApiKey));
            this.searchIndexClient = new SearchIndexClient(
                searchServiceValue,
                CompanyResponseIndexName,
                new SearchCredentials(this.options.SearchServiceQueryApiKey));
            this.searchIndexingIntervalInMinutes = Convert.ToInt32(this.options.SearchIndexingIntervalInMinutes, CultureInfo.InvariantCulture);

            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(this.options.ConnectionString));
            this.companyResponseStorageProvider = companyResponseStorageProvider;
            this.logger = logger;
        }

        /// <summary>
        /// Provide search result for table to be used by SME based on Azure Search service.
        /// </summary>
        /// <param name="searchQuery">Query which the user had typed in messaging extension search field.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <param name="isTaskModuleData">Boolean value which indicates to get the data for company responses task module or not.</param>
        /// <returns>List of search results.</returns>
        public async Task<IList<CompanyResponseEntity>> GetSearchCompanyResponseAsync(string searchQuery, int? count = null, int? skip = null, bool isTaskModuleData = false)
        {
            await this.EnsureInitializedAsync();

            IList<CompanyResponseEntity> companyResponses = new List<CompanyResponseEntity>();
            SearchParameters searchParameters = new SearchParameters
            {
                // Get filtered approved records to show.
                Filter = "ApprovalStatus eq 'Approved'",

                // Get ordered data for company responses to show on messaging extension/task module.
                OrderBy = new[] { "ApprovedOrRejectedDate desc" },
                Top = isTaskModuleData ? count : Constants.DefaultSearchResultCount,
                Skip = skip ?? 0,
                IncludeTotalResultCount = false,
                Select = new[] { "ResponseId", "QuestionLabel", "QuestionText", "ResponseText", "UserId", "CreatedBy", "CreatedDate", "UserRequestType", "LastUpdatedDate", "LastUpdatedBy", "ApproverUserId", "ApprovedOrRejectedBy", "ApprovalStatus", "ApprovalRemark", "ActivityId", "ApprovedOrRejectedDate" },
            };

            var companyResponsesResult = await this.searchIndexClient.Documents.SearchAsync<CompanyResponseEntity>(searchQuery, searchParameters);
            if (companyResponsesResult != null)
            {
                companyResponses = companyResponsesResult.Results.Select(p => p.Document).ToList();
            }

            return companyResponses;
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
                this.searchServiceClient.Dispose();
                this.searchIndexClient.Dispose();
            }

            this.disposed = true;
        }

        /// <summary>
        /// Create index, indexer and data source it doesn't exist.
        /// </summary>
        /// <param name="connectionString">Connection string to the data store.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task InitializeAsync(string connectionString)
        {
            try
            {
                // When there is no company response created and messaging extension is open, table initialization is required here before creating search index or data source or indexer.
                await this.companyResponseStorageProvider.GetCompanyResponsesDataAsync();

                await this.CreateIndexAsync();
                await this.CreateDataSourceAsync(connectionString);
                await this.CreateIndexerAsync();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to initialize Azure Search service: {ex.Message}", SeverityLevel.Error);
                throw;
            }
        }

        /// <summary>
        /// Create index in Azure Search service if it doesn't exist.
        /// </summary>
        /// <returns><see cref="Task"/> That represents index is created if it is not created.</returns>
        private async Task CreateIndexAsync()
        {
            if (!this.searchServiceClient.Indexes.Exists(CompanyResponseIndexName))
            {
                var tableIndex = new Index()
                {
                    Name = CompanyResponseIndexName,
                    Fields = FieldBuilder.BuildForType<CompanyResponseEntity>(),
                };
                await this.searchServiceClient.Indexes.CreateAsync(tableIndex);
            }
        }

        /// <summary>
        /// Add data source if it doesn't exist in Azure Search service.
        /// </summary>
        /// <param name="connectionString">Connection string to the data store.</param>
        /// <returns><see cref="Task"/> That represents data source is added to Azure Search service.</returns>
        private async Task CreateDataSourceAsync(string connectionString)
        {
            if (!this.searchServiceClient.DataSources.Exists(CompanyResponseDataSourceName))
            {
                var dataSource = DataSource.AzureTableStorage(
                                                CompanyResponseDataSourceName,
                                                connectionString,
                                                CompanyResponseTableName);

                await this.searchServiceClient.DataSources.CreateAsync(dataSource);
            }
        }

        /// <summary>
        /// Create indexer if it doesn't exist in Azure Search service.
        /// </summary>
        /// <returns><see cref="Task"/> That represents indexer is created if not available in Azure Search service.</returns>
        private async Task CreateIndexerAsync()
        {
            if (!this.searchServiceClient.Indexers.Exists(CompanyResponseIndexerName))
            {
                var indexer = new Indexer()
                {
                    Name = CompanyResponseIndexerName,
                    DataSourceName = CompanyResponseDataSourceName,
                    TargetIndexName = CompanyResponseIndexName,
                    Schedule = new IndexingSchedule(TimeSpan.FromMinutes(this.searchIndexingIntervalInMinutes)),
                };

                await this.searchServiceClient.Indexers.CreateAsync(indexer);
            }
        }

        /// <summary>
        /// Initialization of InitializeAsync method which will help in indexing.
        /// </summary>
        /// <returns>Task with initialized data.</returns>
        private Task EnsureInitializedAsync()
        {
            return this.initializeTask.Value;
        }
    }
}
