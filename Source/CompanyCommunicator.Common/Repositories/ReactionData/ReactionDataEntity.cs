// <copyright file="ReactionDataEntity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ReactionData
{
    using Microsoft.Azure.Cosmos.Table;

    /// <summary>
    /// Reaction data entity class.
    /// This entity holds the information about a reaction.
    /// </summary>
    public class ReactionDataEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets the reaction id.
        /// </summary>
        public string ReactionId { get; set; }

        /// <summary>
        /// Gets or sets AAD ID of the user.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets conversation ID.
        /// </summary>
        public string ConversationID { get; set; }

        /// <summary>
        /// Gets or sets user's ID.
        /// </summary>
        public string User { get; set; }

        /// <summary>
        /// Gets or sets the reaction.
        /// </summary>
        public string Reaction { get; set; }
    }
}
