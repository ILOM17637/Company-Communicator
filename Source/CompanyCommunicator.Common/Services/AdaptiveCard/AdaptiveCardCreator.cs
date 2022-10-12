// <copyright file="AdaptiveCardCreator.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard
{
    using System;
    using AdaptiveCards;
    using System.Text.Json;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using System.Collections.Generic;
    using Newtonsoft.Json;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;

    /// <summary>
    /// Adaptive Card Creator service.
    /// </summary>
    // Added:subtitle and other features in adaptive card creator file.
    // Added: onbehalfof feature in adapative card creator file
    public class AdaptiveCardCreator
    {
        /// <summary>
        /// Creates an adaptive card.
        /// </summary>
        /// <param name="notificationDataEntity">Notification data entity.</param>
        /// <returns>An adaptive card.</returns>
        public virtual AdaptiveCard CreateAdaptiveCard(
            NotificationDataEntity notificationDataEntity,
            bool voted = false,
            string selectedChoice = "")
        { 
            return this.CreateAdaptiveCard(
                notificationDataEntity.Title,
                notificationDataEntity.Subtitle,
                notificationDataEntity.ImageLink,
                notificationDataEntity.Summary,
                notificationDataEntity.Author,
                notificationDataEntity.ButtonTitle,
                notificationDataEntity.ButtonLink,
                notificationDataEntity.Buttons,
                notificationDataEntity.TrackingUrl,
                notificationDataEntity.ClickRateUrl,
                notificationDataEntity.Id,
                notificationDataEntity.PollOptions,
                notificationDataEntity.PollQuizAnswers,
                notificationDataEntity.IsPollMultipleChoice,
                voted,
                selectedChoice,
                notificationDataEntity.OnBehalfOf);
        }

        /// <summary>
        /// Create an adaptive card instance.
        /// </summary>
        /// <param name="title">The adaptive card's title value.</param>
        /// <param name="subtitle">The adaptive card's subtitle value.</param>
        /// <param name="imageUrl">The adaptive card's image URL.</param>
        /// <param name="summary">The adaptive card's summary value.</param>
        /// <param name="author">The adaptive card's author value.</param>
        /// <param name="buttonTitle">The adaptive card's button title value.</param>
        /// <param name="buttonUrl">The adaptive card's button url value.</param>
        /// <param name="buttons">Buttons.</param>
        /// <param name="trackingurl">Tracking URL for reading purpose.</param>
        /// <param name="clickRateUrl">Click Rate URL for click through rate purpose.</param>
        /// <param name="notificationId">The notification id.</param>
        /// <param name="onBehalfOf">send on behalf of.</param>
        /// <returns>The created adaptive card instance.</returns>
        public AdaptiveCard CreateAdaptiveCard(
            string title,
            string subtitle,
            string imageUrl,
            string summary,
            string author,
            string buttonTitle,
            string buttonUrl,
            string buttons,
            string trackingurl,
            string clickRateUrl,
            string notificationId, 
            string pollOptions,
            string pollQuizAnswers,
            bool isMutipleChoice = false,
            bool voted = false,
            string selectedChoice = "",
            bool onBehalfOf = false)
        {
            var version = new AdaptiveSchemaVersion(1, 0);
            AdaptiveCard card = new AdaptiveCard(version);

            card.Body.Add(new AdaptiveTextBlock()
            {
                Text = title,
                Size = AdaptiveTextSize.ExtraLarge,
                Weight = AdaptiveTextWeight.Bolder,
                Wrap = true,
            });
            if (!string.IsNullOrWhiteSpace(subtitle))
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = subtitle,
                    Wrap = true,
                    Size = AdaptiveTextSize.Large,
                });
            }

            // TODO: increase image width feature
            if (!string.IsNullOrWhiteSpace(imageUrl))
            {
                var img = new AdaptiveImageWithLongUrl()
                {
                    LongUrl = imageUrl,
                    Spacing = AdaptiveSpacing.Default,
                    Size = AdaptiveImageSize.Stretch,
                    AltText = string.Empty,
                };

                // Image enlarge support for Teams web/desktop client.
                img.AdditionalProperties.Add("msteams", new { AllowExpand = true });

                card.Body.Add(img);
            }

            if (!string.IsNullOrWhiteSpace(summary))
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = summary,
                    Wrap = true,
                });
            }

            if (!string.IsNullOrWhiteSpace(author))
            {
                card.Body.Add(new AdaptiveTextBlock()
                {
                    Text = author,
                    Size = AdaptiveTextSize.Small,
                    Weight = AdaptiveTextWeight.Lighter,
                    Wrap = true,
                });
            }

            if (!string.IsNullOrWhiteSpace(pollOptions) && pollOptions != "[]")
            {
                string[] options = JsonConvert.DeserializeObject<string[]>(pollOptions);
                string[] answers = JsonConvert.DeserializeObject<string[]>(pollQuizAnswers);

                var adaptiveCoices = new List<AdaptiveChoice>();
                for (int i = 0; i < options.Length; i++)
                {
                    string optionTitle = options[i];
                    var result = Array.Find(answers, element => element == i.ToString());
                    if (voted && pollQuizAnswers != "[]")
                    {
                        if (!string.IsNullOrWhiteSpace(result))
                        {
                            optionTitle = optionTitle + " " + Strings.PollQuizCorrectAnswer;
                        }
                        else
                        {
                            optionTitle = optionTitle + " " + Strings.PollQuizWrongAnswer;
                        }
                    }

                    adaptiveCoices.Add(new AdaptiveChoice() { Title = optionTitle, Value = i.ToString() });
                }

                var choiceSet = new AdaptiveChoiceSetInput
                {
                    Type = AdaptiveChoiceSetInput.TypeName,
                    Id = "PollChoices",
                    IsRequired = true,
                    ErrorMessage = Strings.PollErrorMessageSelectOption,
                    Style = AdaptiveChoiceInputStyle.Expanded,
                    IsMultiSelect = isMutipleChoice,
                    Choices = adaptiveCoices,
                };

                if (voted)
                {
                    choiceSet.Value = selectedChoice;
                }

                card.Body.Add(choiceSet);

                if (!voted)
                {
                    card.Actions.Add(new AdaptiveSubmitAction()
                    {
                        Title = Strings.PollSubmitVote,
                        Id = "votePoll",
                        Data = "votePoll",
                        DataJson = JsonConvert.SerializeObject(
                        new { notificationId = notificationId }),
                    });
                }
                else
                {
                    if (!string.IsNullOrWhiteSpace(pollQuizAnswers) && pollQuizAnswers != "[]"
                        && !string.IsNullOrWhiteSpace(selectedChoice))
                    {
                        string[] correctAnswers = JsonConvert.DeserializeObject<string[]>(pollQuizAnswers);
                        string[] userAnswers = selectedChoice.Split(',');
                        var set = new HashSet<string>(correctAnswers);
                        bool userFullAnswer = set.SetEquals(userAnswers);

                        if (userFullAnswer)
                        {
                            card.Body.Add(new AdaptiveTextBlock()
                            {
                                Text = Strings.PollQuizCorrect,
                                Size = AdaptiveTextSize.Medium,
                                Weight = AdaptiveTextWeight.Bolder,
                                Color = AdaptiveTextColor.Good,
                                Wrap = true,
                            });
                        }
                        else
                        {
                            card.Body.Add(new AdaptiveTextBlock()
                            {
                                Text = Strings.PollQuizWrong,
                                Size = AdaptiveTextSize.Medium,
                                Weight = AdaptiveTextWeight.Bolder,
                                Color = AdaptiveTextColor.Warning,
                                Wrap = true,
                            });
                        }
                    }
                    else
                    {
                        card.Body.Add(new AdaptiveTextBlock()
                        {
                            Text = Strings.PollThanks,
                            Size = AdaptiveTextSize.Medium,
                            Weight = AdaptiveTextWeight.Bolder,
                            Color = AdaptiveTextColor.Good,
                            Wrap = true,
                        });
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(buttonTitle)
                && !string.IsNullOrWhiteSpace(buttonUrl)
                && string.IsNullOrWhiteSpace(buttons))
            {
                card.Actions.Add(new AdaptiveOpenUrlAction()
                {
                    Title = buttonTitle,
                    Url = new Uri(buttonUrl, UriKind.RelativeOrAbsolute),
                });
            }

            if (!string.IsNullOrWhiteSpace(buttons))
            {
                // enables case insensitive deserialization for card buttons
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                };

                // add the buttons string to the buttons collection for the card
                card.Actions.AddRange(System.Text.Json.JsonSerializer.Deserialize<List<AdaptiveOpenUrlAction>>(buttons, options));
            }

            // ADDED: CLICK THROUGH RATE
            string buttonurl = string.Empty;
            for (var i = 0; i < card.Actions.Count; i++)
            {
                AdaptiveOpenUrlAction action = card.Actions[i] as AdaptiveOpenUrlAction;
                if (action != null)
                {
                    buttonurl = clickRateUrl + "/?url=" + action.Url + "&id=[NotificationID]&userId=[UserID]";
                    action.Url = new Uri(buttonurl, UriKind.RelativeOrAbsolute);
                }
            }

            if (!string.IsNullOrWhiteSpace(trackingurl))
            {
                string trul = trackingurl + "/?id=[ID]&key=[KEY]";

                card.Body.Add(new AdaptiveImage()
                {
                    Url = new Uri(trul, UriKind.RelativeOrAbsolute),
                    Spacing = AdaptiveSpacing.Small,
                    Size = AdaptiveImageSize.Small,
                    IsVisible = false,
                    AltText = string.Empty,
                });
            }

            // TODO: full width adative card feature
            // Full width Adaptive card.
            card.AdditionalProperties.Add("msteams", new { width = "full" });
            return card;
        }
    }
}