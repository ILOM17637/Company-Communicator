// <copyright file="UserTeamsActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Bot
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using AdaptiveCards;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ReactionData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard;
    using Microsoft.Teams.Apps.CompanyCommunicator.Controllers.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Repositories.Extensions;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Company Communicator User Bot.
    /// Captures user data, team data and reaction data.
    /// </summary>

    // Added: Changes related to reactions
    public class UserTeamsActivityHandler : TeamsActivityHandler
    {
        private static readonly string TeamRenamedEventType = "teamRenamed";

        private readonly TeamsDataCapture teamsDataCapture;
        private readonly IBotTelemetryClient botTelemetryClient;
        private readonly IReactionDataRepository reactionDataRepository;
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly AdaptiveCardCreator adaptiveCardCreator;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserTeamsActivityHandler"/> class.
        /// </summary>
        /// <param name="teamsDataCapture">Teams data capture service.</param>
        public UserTeamsActivityHandler(TeamsDataCapture teamsDataCapture, IBotTelemetryClient botTelemetryClient, IReactionDataRepository reactionDataRepository, ISentNotificationDataRepository sentNotificationDataRepository, INotificationDataRepository notificationDataRepository, AdaptiveCardCreator adaptiveCardCreator)
        {
            this.botTelemetryClient = botTelemetryClient ?? throw new ArgumentNullException(nameof(botTelemetryClient));
            this.teamsDataCapture = teamsDataCapture ?? throw new ArgumentNullException(nameof(teamsDataCapture));
            this.reactionDataRepository = reactionDataRepository ?? throw new ArgumentNullException(nameof(reactionDataRepository));
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentNullException(nameof(sentNotificationDataRepository));
            this.notificationDataRepository = notificationDataRepository ??throw new ArgumentNullException(nameof(notificationDataRepository));
            this.adaptiveCardCreator = adaptiveCardCreator ?? throw new ArgumentException(nameof(adaptiveCardCreator));
        }

        protected override async Task OnMessageReactionActivityAsync(ITurnContext<IMessageReactionActivity> turnContext, CancellationToken cancellationToken)
        {
            if (turnContext.Activity.ReactionsAdded != null)
            {
                await this.OnReactionsAddedAsync(turnContext.Activity.ReactionsAdded, turnContext, cancellationToken);
            }

            if (turnContext.Activity.ReactionsRemoved != null)
            {
                await this.OnReactionsRemovedAsync(turnContext.Activity.ReactionsRemoved, turnContext, cancellationToken);
            }
        }

        /// <inheritdoc/>
        protected override async Task OnReactionsAddedAsync(IList<MessageReaction> messageReactions, ITurnContext<IMessageReactionActivity> turnContext, CancellationToken cancellationToken)
        {
            foreach (var reaction in messageReactions)
            {
                // The ReplyToId property of the inbound MessageReaction Activity will correspond to a Message Activity which
                // had previously been sent from this bot.
                var activity = turnContext.Activity;
                if (activity == null)
                {

                    // If we had sent the message from the error handler we wouldn't have recorded the Activity Id and so we
                    // shouldn't expect to see it in the log.
                    throw new NotImplementedException();
                }

                await this.reactionDataRepository.SaveReactionDataAsync(reaction.Type, activity,sentNotificationDataRepository,notificationDataRepository);
            }
        }

        /// <inheritdoc/>
        protected override async Task OnReactionsRemovedAsync(IList<MessageReaction> messageReactions, ITurnContext<IMessageReactionActivity> turnContext, CancellationToken cancellationToken)
        {
            foreach (var reaction in messageReactions)
            {
                // The ReplyToId property of the inbound MessageReaction Activity will correspond to a Message Activity which
                // was previously sent from this bot.
                var activity = turnContext.Activity;
                if (activity == null)
                {
                    // If we had sent the message from the error handler we wouldn't have recorded the Activity Id and so we
                    // shouldn't expect to see it in the log.
                    throw new NotImplementedException();
                }

                await this.reactionDataRepository.RemoveReactionDataAsync(reaction.Type, activity);
            }
        }

        /// <summary>
        /// Invoked when a conversation update activity is received from the channel.
        /// </summary>
        /// <param name="turnContext">The context object for this turn.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects
        /// or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnConversationUpdateActivityAsync(
            ITurnContext<IConversationUpdateActivity> turnContext,
            CancellationToken cancellationToken)
        {
            // base.OnConversationUpdateActivityAsync is useful when it comes to responding to users being added to or removed from the conversation.
            // For example, a bot could respond to a user being added by greeting the user.
            // By default, base.OnConversationUpdateActivityAsync will call <see cref="OnMembersAddedAsync(IList{ChannelAccount}, ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
            // if any users have been added or <see cref="OnMembersRemovedAsync(IList{ChannelAccount}, ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
            // if any users have been removed. base.OnConversationUpdateActivityAsync checks the member ID so that it only responds to updates regarding members other than the bot itself.
            await base.OnConversationUpdateActivityAsync(turnContext, cancellationToken);

            var activity = turnContext.Activity;

            var isTeamRenamed = this.IsTeamInformationUpdated(activity);
            if (isTeamRenamed)
            {
                await this.teamsDataCapture.OnTeamInformationUpdatedAsync(activity);
            }

            if (activity.MembersAdded != null)
            {
                await this.teamsDataCapture.OnBotAddedAsync(turnContext, activity, cancellationToken);
            }

            if (activity.MembersRemoved != null)
            {
                await this.teamsDataCapture.OnBotRemovedAsync(activity);
            }
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            if (!string.IsNullOrEmpty(turnContext.Activity.ReplyToId))
            {
                var txt = turnContext.Activity.Text;
                var value = turnContext.Activity.Value;

                // Check if the activity came from a submit action
                if (string.IsNullOrEmpty(txt) && value != null)
                {
                    var properties2 = new Dictionary<string, string>();
                    properties2.Add("value", turnContext.Activity.Value.ToString());
                    this.LogActivityTelemetry(turnContext.Activity, "TrackValue", properties2);

                    JObject jValue = value as JObject;
                    string notificationId = jValue.ContainsKey("notificationId") ? jValue.Value<string>("notificationId") : string.Empty;

                    string selectedChoice = jValue.ContainsKey("PollChoices") ? jValue.Value<string>("PollChoices") : string.Empty;
                    /*bool? translation = null;
                    if (jValue.ContainsKey("translation"))
                    {
                        translation = jValue.Value<bool>("translation");
                    }*/
                    if (!string.IsNullOrEmpty(selectedChoice))
                    {
                        var notificationEntity2 = await this.notificationDataRepository.GetAsync(NotificationDataTableNames.SentNotificationsPartition, notificationId);
                        var choices = selectedChoice.Split(',');

                        if (notificationEntity2.IsPollQuizMode)
                        {
                            string[] correctAnswers = JsonConvert.DeserializeObject<string[]>(notificationEntity2.PollQuizAnswers);
                            var set = new HashSet<string>(correctAnswers);
                            bool userFullAnswer = set.SetEquals(choices);
                            var quizResult = new Dictionary<string, string>
                            {
                                { "notificationId", notificationId },
                                { "userId", turnContext.Activity.From?.AadObjectId },
                                { "quizResult", userFullAnswer.ToString() },
                            };
                            this.LogActivityTelemetry(turnContext.Activity, "TrackPollQuizResult", quizResult);
                        }

                        foreach (var choice in choices)
                        {
                            var vote = new Dictionary<string, string>
                                {
                                    { "notificationId", notificationId },
                                    { "userId", turnContext.Activity.From?.AadObjectId },
                                    { "vote", choice },
                                };
                            this.LogActivityTelemetry(turnContext.Activity, "TrackPollVote", vote);
                        }

                        // Download base64 data from blob convert to base64 string.
                        if (!string.IsNullOrEmpty(notificationEntity2.ImageBase64BlobName))
                        {
                            notificationEntity2.ImageLink = await this.notificationDataRepository.GetImageAsync(notificationEntity2.ImageLink, notificationEntity2.ImageBase64BlobName);
                        }

                        var card2 = this.adaptiveCardCreator.CreateAdaptiveCard(notificationEntity2, voted: true, selectedChoice: selectedChoice);

                        var adaptiveCardAttachment2 = new Attachment()
                        {
                            ContentType = AdaptiveCard.ContentType,
                            Content = card2,
                        };

                        var newActivity2 = MessageFactory.Attachment(adaptiveCardAttachment2);
                        newActivity2.Id = turnContext.Activity.ReplyToId;
                        newActivity2.Summary = notificationEntity2.Title;
                        await turnContext.UpdateActivityAsync(newActivity2, cancellationToken);
                    }
                }
            }
        }

        private void LogActivityTelemetry(IActivity activity, string eventName = "UserInfo", Dictionary<string, string> properties = null)
        {
            if (properties == null)
            {
                properties = new Dictionary<string, string>();
            }

            properties.Add("UserAadObjectId", activity.From?.AadObjectId);
            properties.Add("UserName", activity.From?.Name);

            // client info
            var clientInfoEntity = activity.Entities?.Where(e => e.Type == "clientInfo")?.FirstOrDefault();
            properties.Add("Locale", clientInfoEntity?.Properties["locale"]?.ToString());
            properties.Add("Country", clientInfoEntity?.Properties["country"]?.ToString());
            properties.Add("TimeZone", clientInfoEntity?.Properties["timezone"]?.ToString());
            properties.Add("Platform", clientInfoEntity?.Properties["platform"]?.ToString());

            properties.Add("ActivityId", activity.Id);
            properties.Add("ActivityType", activity.Type);
            properties.Add("ConversationType", string.IsNullOrWhiteSpace(activity.Conversation?.ConversationType) ? "personal" : activity.Conversation.ConversationType);
            properties.Add("ConversationId", activity.Conversation?.Id);

            var channelData = activity.GetChannelData<TeamsChannelData>();
            properties.Add("TeamId", channelData?.Team?.Id);

            this.botTelemetryClient.TrackEvent(eventName, properties);
        }

        private bool IsTeamInformationUpdated(IConversationUpdateActivity activity)
        {
            if (activity == null)
            {
                return false;
            }

            var channelData = activity.GetChannelData<TeamsChannelData>();
            if (channelData == null)
            {
                return false;
            }

            return UserTeamsActivityHandler.TeamRenamedEventType.Equals(channelData.EventType, StringComparison.OrdinalIgnoreCase);
        }
    }
}