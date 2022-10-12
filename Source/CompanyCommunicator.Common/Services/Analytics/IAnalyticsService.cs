namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Analytics
{
    using System.Threading;
    using System.Threading.Tasks;

    public interface IAnalyticsService
    {
        Task<KustoQueryResult> GetPollVoteResultByNotificationIdAsync(string notificationId, CancellationToken cancellationToken = default(CancellationToken));

        Task<int> GetFullyCorrectQuizAnswersCountByNotificationIdAsync(string notificationId, CancellationToken cancellationToken = default(CancellationToken));

        Task<int> GetUniquePollVotesCountByNotificationIdAsync(string notificationId, CancellationToken cancellationToken = default(CancellationToken));
    }
}