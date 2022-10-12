using Microsoft.AspNetCore.Mvc;
using Microsoft.Bot.Builder;
using Microsoft.Extensions.Logging;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    /// <summary>
    /// Controller for button redirects.
    /// </summary>
    [Route("redirect")]
    [ApiController]
    public class RedirectController : ControllerBase
    {
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly INotificationDataRepository notificationDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="RedirectController"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">Notification Data Repository.</param>
        /// <param name="sentNotificationDataRepository">sent Notification Data Repository.</param>
        public RedirectController(INotificationDataRepository notificationDataRepository, ISentNotificationDataRepository sentNotificationDataRepository)
        {
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentNullException(nameof(sentNotificationDataRepository));
        }

        [HttpGet]
        public async Task<IActionResult> Get(string url, string id, string userId)
        {
            // url cannot be null
            if (string.IsNullOrWhiteSpace(url))
            {
                throw new ArgumentNullException(nameof(url));
            }

            // id cannot be null
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentNullException(nameof(id));
            }

            // userId cannot be null
            if (string.IsNullOrWhiteSpace(userId))
            {
                throw new ArgumentNullException(nameof(userId));
            }

            // gets the sent notification object for the message sent
            var sentnotificationEntity = await this.sentNotificationDataRepository.GetAsync(id, userId);

            // if we have a instance that was sent to a user
            if (sentnotificationEntity != null)
            {
                // if the message was not read yet
                if (sentnotificationEntity.ClickStatus != true)
                {
                    sentnotificationEntity.ClickStatus = true;
                    sentnotificationEntity.ClickDate = DateTime.UtcNow;

                    await this.sentNotificationDataRepository.CreateOrUpdateAsync(sentnotificationEntity);

                    // gets the sent notification summary that needs to be updated
                    var notificationEntity = await this.notificationDataRepository.GetAsync(
                        NotificationDataTableNames.SentNotificationsPartition,
                        id);

                    // if the notification entity is null it means it doesnt exist or is not a sent message yet
                    if (notificationEntity != null)
                    {
                        // increment the number of reads
                        notificationEntity.Clicks++;

                        // persists the change
                        await this.notificationDataRepository.CreateOrUpdateAsync(notificationEntity);
                    }
                }
            }

            return this.Redirect(url);
        }
    }
}
