// <copyright file="BaseNotification.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    using System;

    /// <summary>
    /// Base notification model class.
    /// </summary>
    public class BaseNotification
    {
        /// <summary>
        /// Gets or sets Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets Title value.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets Subtitle value.
        /// </summary>
        public string Subtitle { get; set; }
        /// <summary>
        /// Gets or sets the Image Link value.
        /// </summary>
        public string ImageLink { get; set; }

        /// <summary>
        /// Gets or sets the blob name for the image in base64 format.
        /// </summary>
        public string ImageBase64BlobName { get; set; }

        /// <summary>
        /// Gets or sets the Summary value.
        /// </summary>
        public string Summary { get; set; }

        /// <summary>
        /// Gets or sets the Author value.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets the Button Title value.
        /// </summary>
        public string ButtonTitle { get; set; }

        /// <summary>
        /// Gets or sets the Button Link value.
        /// </summary>
        public string ButtonLink { get; set; }

        /// <summary>
        /// Gets or sets the Buttons value.
        /// </summary>
        public string Buttons { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether gets or sets the IsScheduled value.
        /// </summary>
        public bool IsScheduled { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether gets or sets the IsImportant value.
        /// </summary>
        public bool IsImportant { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether OnBehalfOf set to true or false.
        /// </summary>
        public bool OnBehalfOf { get; set; }

        /// <summary>
        /// Gets or sets the Created DateTime value.
        /// </summary>
        public DateTime CreatedDateTime { get; set; }

        /// <summary>
        /// Gets or sets poll options.
        /// </summary>
        public string PollOptions { get; set; }

        /// <summary>
        /// Gets or sets the messge type.
        /// </summary>
        public string MessageType { get; set; }

        /// <summary>
        /// Gets or sets poll quiz answers.
        /// </summary>
        public string PollQuizAnswers { get; set; }

        /// <summary>
        /// Gets or sets is Poll Quiz mode or not.
        /// </summary>
        public bool IsPollQuizMode { get; set; }

        /// <summary>
        /// Gets or sets whether is Poll MCQ or not.
        /// </summary>
        public bool IsPollMultipleChoice { get; set; }

    }
}
