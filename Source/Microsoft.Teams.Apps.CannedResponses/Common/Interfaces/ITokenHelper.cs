// <copyright file="ITokenHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Common.Interfaces
{
    using System;

    /// <summary>
    /// Interface for token helper class which helps in generating custom JWT token and retrieval of user access token.
    /// </summary>
    public interface ITokenHelper
    {
        /// <summary>
        /// Generate JWT token used by client app to authenticate HTTP calls with API.
        /// </summary>
        /// <param name="serviceURL">Service URL from bot.</param>
        /// <param name="userObjectIdentifier">Aad object id of user.</param>
        /// <param name="jwtExpiryMinutes">Expiry of token in minutes.</param>
        /// <returns>JWT token.</returns>
        string GenerateInternalAPIToken(Uri serviceURL, string userObjectIdentifier, int jwtExpiryMinutes);
    }
}
