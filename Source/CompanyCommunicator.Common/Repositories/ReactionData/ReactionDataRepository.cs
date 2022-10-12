// <copyright file="ReactionDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ReactionData
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;

    /// <summary>
    /// Repository of the reaction data stored in the table storage.
    /// </summary>
    public class ReactionDataRepository : BaseRepository<ReactionDataEntity>, IReactionDataRepository
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ReactionDataRepository"/> class.
        /// </summary>
        /// <param name="logger">The logging service.</param>
        /// <param name="repositoryOptions">Options used to create the repository.</param>
        public ReactionDataRepository(
            ILogger<ReactionDataRepository> logger,
            IOptions<RepositoryOptions> repositoryOptions)
            : base(
                  logger,
                  storageAccountConnectionString: repositoryOptions.Value.StorageAccountConnectionString,
                  tableName: ReactionDataTableNames.TableName,
                  defaultPartitionKey: ReactionDataTableNames.ReactionDataPartition,
                  ensureTableExists: repositoryOptions.Value.EnsureTableExists)
        {
        }

        /// <inheritdoc/>
        public async Task EnsureReactionDataTableExistsAsync()
        {
            var exists = await this.Table.ExistsAsync();
            if (!exists)
            {
                await this.Table.CreateAsync();
            }
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<ReactionDataEntity>> GetReactionDataEntitiesByIdsAsync(IEnumerable<string> reactionIds)
        {
            var rowKeysFilter = this.GetRowKeysFilter(reactionIds);

            return await this.GetWithFilterAsync(rowKeysFilter);
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<string>> GetReactionNamesByIdsAsync(IEnumerable<string> ids)
        {
            if (ids == null || !ids.Any())
            {
                return new List<string>();
            }

            var rowKeysFilter = this.GetRowKeysFilter(ids);
            var reactionDataEntities = await this.GetWithFilterAsync(rowKeysFilter);

            return reactionDataEntities.Select(p => p.Name).OrderBy(p => p);
        }

        /// <inheritdoc/>
        public async Task<IEnumerable<ReactionDataEntity>> GetAllSortedAlphabeticallyByNameAsync()
        {
            var reactionDataEntities = await this.GetAllAsync();
            var sortedSet = new SortedSet<ReactionDataEntity>(reactionDataEntities, new ReactionDataEntityComparer());
            return sortedSet;
        }

        private class ReactionDataEntityComparer : IComparer<ReactionDataEntity>
        {
            public int Compare(ReactionDataEntity x, ReactionDataEntity y)
            {
                return x.Name.CompareTo(y.Name);
            }
        }
    }
}
