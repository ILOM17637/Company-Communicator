// <copyright file="ReactionData.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>
namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.Export.Model
{
    /// <summary>
    /// the model class for team data.
    /// </summary>
    public class ReactionData
    {
        /// <summary>
        /// Gets or sets the Reaction id value.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets the User Name id value.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the Reaction value.
        /// </summary>
        public string Reaction { get; set; }

        /// <summary>
        /// Gets or sets the Conversation id value.
        /// </summary>
        public string ConversationId { get; set; }
    }
}
