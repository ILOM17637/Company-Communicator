// <copyright file="ReactionDataTableNames.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ReactionData
{
    /// <summary>
    /// Reaction data table names.
    /// </summary>
    public static class ReactionDataTableNames
    {
        /// <summary>
        /// Table name for the reaction data table.
        /// </summary>
        public static readonly string TableName = "ReactionData";

        /// <summary>
        /// Reaction data partition key name.
        /// </summary>
        public static readonly string ReactionDataPartition = "ReactionData";
    }
}
