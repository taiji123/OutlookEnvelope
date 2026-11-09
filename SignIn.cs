using System;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Win32;

internal static class Settings
{
    const string Key = @"Software\OutlookEnvelope";
    public static object Get(string name) => Registry.GetValue(@"HKEY_CURRENT_USER\" + Key, name, null);
    public static void Set(string name, object value)
    {
        using (var key = Registry.CurrentUser.CreateSubKey(Key))
        {
            if (value == null) key.DeleteValue(name, false);
            else key.SetValue(name, value);
        }
    }
}

internal sealed class SignIn
{
    readonly IPublicClientApplication client;
    static readonly string[] Scopes = { "https://graph.microsoft.com/Mail.ReadBasic" };
    readonly string cacheName;
    public SignIn(Guid tenant, Guid app)
    {
        cacheName = "Session-" + tenant + "-" + app;
        client = PublicClientApplicationBuilder.Create(app.ToString())
            .WithAuthority("https://login.microsoftonline.com/" + tenant)
            .WithRedirectUri("http://localhost").Build();
        client.UserTokenCache.SetBeforeAccess(args =>
        {
            if (Settings.Get(cacheName) is byte[] data)
                args.TokenCache.DeserializeMsalV3(ProtectedData.Unprotect(data, null, DataProtectionScope.CurrentUser));
        });
        client.UserTokenCache.SetAfterAccess(args =>
        {
            if (args.HasStateChanged) Settings.Set(cacheName,
                ProtectedData.Protect(args.TokenCache.SerializeMsalV3(), null, DataProtectionScope.CurrentUser));
        });
    }
    public async Task<AuthenticationResult> Get(bool interactive, bool refresh, CancellationToken cancel)
    {
        if (interactive) return await client.AcquireTokenInteractive(Scopes)
            .WithUseEmbeddedWebView(false).WithPrompt(Prompt.SelectAccount).ExecuteAsync(cancel);
        var account = (await client.GetAccountsAsync()).FirstOrDefault(a => a.HomeAccountId.Identifier == (string)Settings.Get("Account"));
        return await client.AcquireTokenSilent(Scopes, account).WithForceRefresh(refresh).ExecuteAsync(cancel);
    }
}
