// <copyright file="BaseStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.Providers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.RetryPolicies;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which initializes table if not exists and provide table client instance.
    /// </summary>
    public class BaseStorageProvider
    {
        /// <summary>
        /// Microsoft Azure Table storage connection string.
        /// </summary>
        private readonly string connectionString;

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseStorageProvider"/> class.
        /// Handles Microsoft Azure Table creation.
        /// </summary>
        /// <param name="connectionString">Azure Table storage connection string.</param>
        /// <param name="tableName">Azure Table storage table name.</param>
        public BaseStorageProvider(string connectionString, string tableName)
        {
            this.InitializeTask = new Lazy<Task>(() => this.InitializeAsync());
            this.connectionString = connectionString ?? throw new ArgumentNullException(nameof(connectionString));
            this.TableName = tableName;
        }

        /// <summary>
        /// Gets or sets task for initialization.
        /// </summary>
        protected Lazy<Task> InitializeTask { get; set; }

        /// <summary>
        /// Gets or sets Microsoft Azure Table storage table name.
        /// </summary>
        protected string TableName { get; set; }

        /// <summary>
        /// Gets or sets a table in the Microsoft Azure Table storage.
        /// </summary>
        protected CloudTable ResponsesCloudTable { get; set; }

        /// <summary>
        /// Ensures Microsoft Azure Table storage should be created before working on table.
        /// </summary>
        /// <returns>Represents an asynchronous operation.</returns>
        protected async Task EnsureInitializedAsync()
        {
            await this.InitializeTask.Value;
        }

        /// <summary>
        /// Create tables if it doesn't exist.
        /// </summary>
        /// <returns>Asynchronous task which represents table is created if its not existing.</returns>
        private async Task InitializeAsync()
        {
            // Exponential retry policy with backoff of 3 seconds and 5 retries.
            var exponentialRetryPolicy = new TableRequestOptions()
            {
                RetryPolicy = new ExponentialRetry(TimeSpan.FromSeconds(3), 5),
            };

            try
            {
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(this.connectionString);
                CloudTableClient cloudTableClient = storageAccount.CreateCloudTableClient();
                cloudTableClient.DefaultRequestOptions = exponentialRetryPolicy;
                this.ResponsesCloudTable = cloudTableClient.GetTableReference(this.TableName);
                await this.ResponsesCloudTable.CreateIfNotExistsAsync();
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
