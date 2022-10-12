
namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;


    /// <summary>
    /// Controller to get app settings.
    /// </summary>
    [Route("api/settings")]
    [ApiController]
    public class AppSettingsController : ControllerBase
    {
        private readonly BotOptions botOptions;

        /// <summary>
        /// Initializes a new instance of the <see cref="AppSettingsController"/> class.
        /// </summary>
        /// <param name="userAppOptions">User app options.</param>
        public AppSettingsController(IOptions<BotOptions> userAppOptions)
        {
            this.botOptions = userAppOptions?.Value ?? throw new ArgumentNullException(nameof(userAppOptions));
        }

        /// <summary>
        /// Get value for ImageUploadBlobStorage.
        /// </summary>
        /// <returns>Required sent notification.</returns>
        [HttpGet]
        public IActionResult GetAppSettings()
        {
            var appId = this.botOptions.AuthorAppId;
            var imageUploadBlobStorage = this.botOptions.ImageUploadBlobStorage;
            var response = new AppConfigurations()
            {
                ImageUploadBlobStorage = imageUploadBlobStorage,
            };

            return this.Ok(response);
        }
    }
}
