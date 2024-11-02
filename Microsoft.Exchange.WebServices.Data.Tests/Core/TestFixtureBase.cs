using System.Net;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace Microsoft.Exchange.WebServices.Data.Tests.Core;

public class TestFixtureBase
{
    private ServiceProvider ServiceProvider { get; set; }
    
    protected ExchangeService GetExchangeService(string? testUser = null)
    {
        var settings = ServiceProvider.GetRequiredService<ApplicationSettings>();
        var service = new ExchangeService(GetWorkaroundTimeZone())
        {
            Url = new Uri(settings.EwsServiceUrl),
            Credentials = new WebCredentials(settings.Username, settings.Password),
            ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, testUser ?? TestUsers.User1)
        };
        
        return service;
    }
    
    private static TimeZoneInfo GetWorkaroundTimeZone()
    {
        // https://stackoverflow.com/questions/39467609/the-specified-time-zone-isnt-valid-using-ews-from-server
        return TimeZoneInfo.CreateCustomTimeZone(
            id: "Time zone to workaround a bug",
            baseUtcOffset: TimeZoneInfo.Local.BaseUtcOffset,
            displayName: "Time zone to workaround a bug", "Time zone to workaround a bug"
        );
    }
    
    [OneTimeSetUp]
    public void OneTimeSetUp()
    {
        // Settings
        var applicationSettings = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json")
            .AddEnvironmentVariables()
            .AddUserSecrets<ApplicationSettings>()
            .Build()
            .Get<ApplicationSettings>();
        if (applicationSettings == null)
        {
            throw new ArgumentException("Application settings not found");
        }
        
        // Services
        var services = new ServiceCollection()
            .AddSingleton(applicationSettings);
        
        // Infrastructure
        ServicePointManager.ServerCertificateValidationCallback 
            += (sender, certificate, chain, sslPolicyErrors) => true;
        
        // Build service provider
        ServiceProvider = services
            .BuildServiceProvider();
    }

    [OneTimeTearDown]
    public void OneTimeTearDown()
    {
        ServiceProvider.Dispose();
    }
}