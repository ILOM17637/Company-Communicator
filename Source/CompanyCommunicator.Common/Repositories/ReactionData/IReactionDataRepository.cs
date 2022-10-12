// <copyright file="IReactionDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>
namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ReactionData
{
    using System.Collections.Generic;
    using System.Threading.Tasks;

    /// <summary>
    /// Interface for Reaction Data Repository.
    /// </summary>
    public interface IReactionDataRepository : IRepository<ReactionDataEntity>
    {
        /// <summary>
        /// This method ensures the ExportData table is created in the storage.
        /// This method should be called before kicking off an Azure function that uses the ExportData table.
        /// Otherwise the app will crash.
        /// By design, Azure functions (in this app) do not create a table if it's absent.
        /// </summary>
        /// <returns>A <see cref="Task"/> representing the result of the asynchronous operation.</returns>
        public Task EnsureReactionDataTableExistsAsync();

        /// <summary>
        /// Gets reaction data entities by ID values.
        /// </summary>
        /// <param name="reactionIds">Reaction IDs.</param>
        /// <returns>Reaction data entities.</returns>
        public Task<IEnumerable<ReactionDataEntity>> GetReactionDataEntitiesByIdsAsync(IEnumerable<string> reactionIds);

        /// <summary>
        /// Get reaction names by Ids.
        /// </summary>
        /// <param name="ids">Reaction ids.</param>
        /// <returns>Names of the reactions matching incoming ids.</returns>
        public Task<IEnumerable<string>> GetReactionNamesByIdsAsync(IEnumerable<string> ids);

        /// <summary>
        /// Get all reaction data entities, and sort the result alphabetically by name.
        /// </summary>
        /// <returns>The reaction data entities sorted alphabetically by name.</returns>
        public Task<IEnumerable<ReactionDataEntity>> GetAllSortedAlphabeticallyByNameAsync();
    }
}
