using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
class Constants
{
    public const string Tenant = "e4a09b27-9de1-4dba-a0bf-84c861f3fe2e";

    // Admin consent flow
    public const string AuthorityUri = "https://login.microsoftonline.com/" + Tenant;
    public const string RedirectUriForAppAuthn = "https://login.microsoftonline.com";
    public const string ClientIdForAppAuthn = "06b7c8b0-433b-4b64-ad4f-8f076f9d14e9";
    public const string ClientSecret = "kUn4w7RgeS/4=:mL+AUUn55J5Gu.6fmz";
}