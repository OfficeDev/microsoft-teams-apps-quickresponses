// <copyright file="TokenHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CannedResponses.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.IdentityModel.Tokens.Jwt;
    using System.Security.Claims;
    using System.Text;
    using Microsoft.Extensions.Options;
    using Microsoft.IdentityModel.Tokens;
    using Microsoft.Teams.Apps.CannedResponses.Common.Interfaces;
    using Microsoft.Teams.Apps.CannedResponses.Models;

    /// <summary>
    /// Helper class for JWT token generation and validation.
    /// </summary>
    public class TokenHelper : ITokenHelper
    {
        /// <summary>
        /// Represents service URL string.
        /// </summary>
        private const string ServiceURL = "serviceURL";

        /// <summary>
        /// Represents user object identifier string.
        /// </summary>
        private const string UserObjectIdentifier = "userObjectIdentifier";

        /// <summary>
        /// Retrieve required bot configurations.
        /// </summary>
        private readonly IOptions<BotSetting> options;

        /// <summary>
        /// Initializes a new instance of the <see cref="TokenHelper"/> class.
        /// </summary>
        /// <param name="options">A set of key/value application configuration properties for bot.</param>
        public TokenHelper(IOptions<BotSetting> options)
        {
            this.options = options;
        }

        /// <summary>
        /// Generate JWT token used by client application to authenticate HTTP calls with API.
        /// </summary>
        /// <param name="serviceURL">Service URL from bot.</param>
        /// <param name="userObjectIdentifier">Aad object id of user.</param>
        /// <param name="jwtExpiryMinutes">Expiry of token in minutes.</param>
        /// <returns>JWT token.</returns>
        public string GenerateInternalAPIToken(Uri serviceURL, string userObjectIdentifier, int jwtExpiryMinutes)
        {
            if (serviceURL == null)
            {
                throw new ArgumentNullException(nameof(serviceURL));
            }

            SymmetricSecurityKey signingKey = new SymmetricSecurityKey(Encoding.ASCII.GetBytes(this.options.Value.SecurityKey));
            SigningCredentials signingCredentials = new SigningCredentials(signingKey, SecurityAlgorithms.HmacSha256);

            SecurityTokenDescriptor securityTokenDescriptor = new SecurityTokenDescriptor()
            {
                Subject = new ClaimsIdentity(
                    new List<Claim>()
                    {
                        new Claim(ServiceURL, serviceURL.ToString()),
                        new Claim(UserObjectIdentifier, userObjectIdentifier),
                    }, "Custom"),
                NotBefore = DateTime.UtcNow,
                SigningCredentials = signingCredentials,
                Issuer = this.options.Value.AppBaseUri,
                Audience = this.options.Value.AppBaseUri,
                IssuedAt = DateTime.UtcNow,
                Expires = DateTime.UtcNow.AddMinutes(jwtExpiryMinutes),
            };

            JwtSecurityTokenHandler tokenHandler = new JwtSecurityTokenHandler();
            SecurityToken token = tokenHandler.CreateToken(securityTokenDescriptor);
            return tokenHandler.WriteToken(token);
        }
    }
}
