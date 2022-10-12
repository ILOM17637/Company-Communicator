// <copyright file="ReactionDataMap.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

// Added: Reaction Data map file
namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.Export.Mappers
{
    using System;
    using CsvHelper.Configuration;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;
    using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.Export.Model;

    /// <summary>
    /// Mapper class for Reaction.
    /// </summary>
    public sealed class ReactionDataMap : ClassMap<ReactionData>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ReactionDataMap"/> class.
        /// </summary>
        public ReactionDataMap()
        {
            this.Map(x => x.Id).Name("Reaction_Id");
            this.Map(x => x.Name).Name("UPN");
            this.Map(x => x.Reaction).Name("Reaction");
            this.Map(x => x.ConversationId).Name("Conversation_Id");
        }
    }
}