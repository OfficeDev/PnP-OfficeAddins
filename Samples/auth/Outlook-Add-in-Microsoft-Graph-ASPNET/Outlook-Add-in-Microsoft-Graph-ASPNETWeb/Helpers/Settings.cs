// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using System;
using System.Configuration;
using System.Web;

namespace OutlookAddinMicrosoftGraphASPNET.Helpers
{
    /// <summary>
    /// Provides management of basic user and web application authentication and authorization information. 
    /// </summary>
    public static class Settings
    {
        public static string AzureADClientId = ConfigurationManager.AppSettings["AAD:ClientID"];
        public static string AzureADClientSecret = ConfigurationManager.AppSettings["AAD:ClientSecret"];

        public static string AzureADAuthority = @"https://login.microsoftonline.com/" + ConfigurationManager.AppSettings["AAD:O365TenantID"] + "/oauth2/v2.0";
        public static string AzureADLogoutAuthority = @"https://login.microsoftonline.com/common/oauth2/logout?post_logout_redirect_uri=";
        public static string GraphApiResource = @"https://graph.microsoft.com/";

       
        /// <summary>
        /// Ensures that the current key to the SessionToken table is in the cookie.
        /// </summary>
        /// <param name="ctx">The HTTP request context.</param>
        /// <returns></returns>
        public static string GetUserAuthStateId(HttpContextBase ctx)
        {
            string id;
            if (ctx.Request.Cookies[SessionKeys.Login.UserAuthStateId] == null)
            {
                // Convert GUID to a string and format as numeral to remove hyphens.
                id = Guid.NewGuid().ToString("N");
                ctx.Response.Cookies.Add(new HttpCookie(SessionKeys.Login.UserAuthStateId)
                {
                    Expires = DateTime.Now.AddMinutes(20),
                    Value = id,
                    Secure = true,
                    SameSite = SameSiteMode.None
                });
            }
            else
            {
                id = ctx.Request.Cookies[SessionKeys.Login.UserAuthStateId].Value;
            }

            return id;
        }
    }
}
