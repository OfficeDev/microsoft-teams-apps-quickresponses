// <copyright file="ResourceController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Controllers
{
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Localization;

    /// <summary>
    /// Controller to provide localized resource strings.
    /// </summary>
    [Route("api/Resource")]
    [ApiController]
    [Authorize]

    public class ResourceController : ControllerBase
    {
        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Initializes a new instance of the <see cref="ResourceController"/> class.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        public ResourceController(IStringLocalizer<Strings> localizer)
        {
            this.localizer = localizer;
        }

        /// <summary>
        /// Get localized resource strings.
        /// </summary>
        /// <returns>Object containing resource strings.</returns>
        [HttpGet]
        [Route("ResourceStrings")]
        public IActionResult GetResourceStrings()
        {
            var resourceStrings = new
            {
                Label = this.localizer.GetString("Label").Value,
                TypeCategoryPlaceholder = this.localizer.GetString("TypeCategoryPlaceholder").Value,
                TypeQuestionPlaceholder = this.localizer.GetString("TypeQuestionPlaceholder").Value,
                TypeResponsePlaceholder = this.localizer.GetString("TypeResponsePlaceholder").Value,
                Questions = this.localizer.GetString("Questions").Value,
                Response = this.localizer.GetString("Response").Value,
                BackButtonText = this.localizer.GetString("BackButtonText").Value,
                SuggestButtonText = this.localizer.GetString("SuggestButtonText").Value,
                FieldRequiredMessage = this.localizer.GetString("FieldRequiredMessage").Value,
                MyResponsesText = this.localizer.GetString("MyResponsesText").Value,
                AddNewButtonText = this.localizer.GetString("AddNewButtonText").Value,
                EditButtonText = this.localizer.GetString("EditButtonText").Value,
                CancelButtonText = this.localizer.GetString("CancelButtonText").Value,
                ConfirmButtonText = this.localizer.GetString("ConfirmButtonText").Value,
                DialogConfirmText = this.localizer.GetString("DialogConfirmText").Value,
                DialogConfirmHeader = this.localizer.GetString("DialogConfirmHeader").Value,
                DeleteButtonText = this.localizer.GetString("DeleteButtonText").Value,
                SearchPlaceholder = this.localizer.GetString("SearchPlaceholder").Value,
                Question = this.localizer.GetString("Question").Value,
                CompanyResponsesMenuText = this.localizer.GetString("CompanyResponsesMenuText").Value,
                RequestsMenuText = this.localizer.GetString("RequestsMenuText").Value,
                RequestedOnText = this.localizer.GetString("RequestedOnText").Value,
                AddButtonText = this.localizer.GetString("AddButtonText").Value,
                UpdateAndShareButtonText = this.localizer.GetString("UpdateAndShareButtonText").Value,
                UpdateButtonText = this.localizer.GetString("UpdateButtonText").Value,
                MaxResponsesMessage = this.localizer.GetString("MaxResponsesMessage").Value,
                StatusText = this.localizer.GetString("StatusText").Value,
                SuggestNewButtonText = this.localizer.GetString("SuggestNewButtonText").Value,
                MaxCompanyResponseMessage = this.localizer.GetString("MaxCompanyResponseMessage").Value,
            };

            return this.Ok(resourceStrings);
        }

        /// <summary>
        /// Get localized error resource strings.
        /// </summary>
        /// <returns>Object containing resource strings.</returns>
        [HttpGet]
        [Route("ErrorResourceStrings")]
        public IActionResult GetErrorResourceStrings()
        {
            var resourceStrings = new
            {
                UnauthorizedErrorMessage = this.localizer.GetString("UnauthorizedErrorMessage").Value,
                ForbiddenErrorMessage = this.localizer.GetString("ForbiddenErrorMessage").Value,
                GeneralErrorMessage = this.localizer.GetString("GeneralErrorMessage").Value,
                RefreshLinkText = this.localizer.GetString("RefreshLinkText").Value,
            };

            return this.Ok(resourceStrings);
        }
    }
}