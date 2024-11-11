using System.Net;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace Microsoft.Exchange.WebServices.Data.Tests.Core;

public class TestFixtureBase
{
    private ServiceProvider ServiceProvider { get; set; }
    
    protected ExchangeService GetExchangeServiceUsingImpersonation(string? testUser = null)
    {
        var settings = ServiceProvider.GetRequiredService<ApplicationSettings>();
        var service = new ExchangeService(GetWorkaroundTimeZone())
        {
            Url = new Uri(settings.EwsServiceUrl),
            Credentials = new WebCredentials(settings.Impersonation.Username, settings.Impersonation.Password),
            ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, testUser ?? TestUsers.User1)
        };
        
        return service;
    }
    
    protected ExchangeService GetExchangeServiceUsingDirectAccess()
    {
        var settings = ServiceProvider.GetRequiredService<ApplicationSettings>();
        var service = new ExchangeService(GetWorkaroundTimeZone())
        {
            Url = new Uri(settings.EwsServiceUrl),
            Credentials = new WebCredentials(settings.DirectAccess.Username, settings.DirectAccess.Password),
        };
        
        return service;
    }
    
    protected ExchangeService GetExchangeServiceUsingDelegatingAccess()
    {
        var settings = ServiceProvider.GetRequiredService<ApplicationSettings>();
        var service = new ExchangeService(GetWorkaroundTimeZone())
        {
            Url = new Uri(settings.EwsServiceUrl),
            Credentials = new WebCredentials(settings.DelegatingAccess.Username, settings.DelegatingAccess.Password),
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

    private void DeleteExistingAppointments(string testUser)
    {
        var exchangeService = GetExchangeServiceUsingImpersonation(testUser);
        
        var calendar = CalendarFolder.Bind(exchangeService, WellKnownFolderName.Calendar, []);
        var calendarView = new CalendarView(DateTime.Now.AddMonths(-1), DateTime.Now.AddMonths(+1), int.MaxValue)
        {
            PropertySet = new PropertySet(ItemSchema.Id)
        };
        var appointments = calendar.FindAppointments(calendarView).ToArray();

        if (appointments.Length > 0)
        {
            exchangeService.DeleteItems(
                appointments.Select(a => a.Id).ToArray(),
                DeleteMode.HardDelete,
                SendCancellationsMode.SendToNone,
                AffectedTaskOccurrence.AllOccurrences
            );
        }
    }

    [OneTimeTearDown]
    public void OneTimeTearDown()
    {
        // Clean up data
        DeleteExistingAppointments(TestUsers.User1);
        DeleteExistingAppointments(TestUsers.User2);
        DeleteExistingAppointments(TestUsers.User3);
        DeleteExistingAppointments(TestUsers.User4);
        DeleteExistingAppointments(TestUsers.User5);
        
        // Dispose resources
        ServiceProvider.Dispose();
    }
}