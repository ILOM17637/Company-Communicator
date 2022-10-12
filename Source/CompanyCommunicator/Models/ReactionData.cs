// <copyright file="ReactionData.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    /// <summary>
    /// User's Reaction summary model class.
    /// </summary>
    public class ReactionData
    {
        /// <summary>
        /// Gets or sets message Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets user's name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets user's Emotion.
        /// </summary>
        public string Reaction { get; set; }
    }
}