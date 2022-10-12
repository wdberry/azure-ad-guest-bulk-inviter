using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Microsoft.Authentication.WebAssembly.Msal;
using Microsoft.AspNetCore.Components.WebAssembly.Authentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using azure_ad_guest_bulk_inviter;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });
builder.Services.AddMsalAuthentication<RemoteAuthenticationState, RemoteUserAccount>(options =>
{
    builder.Configuration.Bind("AzureAd", options.ProviderOptions.Authentication);

    options.ProviderOptions.DefaultAccessTokenScopes.Add(
            "https://graph.microsoft.com/User.Read");
    //options.ProviderOptions.AdditionalScopesToConsent.Add("https://graph.microsoft.com/User.Invite.All");
    //options.ProviderOptions.AdditionalScopesToConsent.Add("https://graph.microsoft.com/User.ReadWrite.All");
});
builder.Services.AddGraphClient("https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/User.Invite.All");
//builder.Services.AddGraphClient("https://graph.microsoft.com/User.Read");

await builder.Build().RunAsync();
