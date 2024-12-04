using System.Net;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace Microsoft.Exchange.WebServices.Data.Tests.Core;

public class TestFixtureBase
{
    private ServiceProvider ServiceProvider { get; set; }
    protected TestSettings Settings => ServiceProvider.GetRequiredService<TestSettings>();
    
    protected ExchangeService GetExchangeServiceUsingDirectAccess(UserCredentials userCredentials)
    {
        var service = new ExchangeService(GetWorkaroundTimeZone())
        {
            Url = new Uri(Settings.EwsServiceUrl),
            Credentials = new WebCredentials(userCredentials.Username, userCredentials.Password),
        };
        
        return service;
    }
    
    protected ExchangeService GetExchangeServiceUsingImpersonation(UserCredentials userCredentials)
    {
        var service = new ExchangeService(GetWorkaroundTimeZone())
        {
            Url = new Uri(Settings.EwsServiceUrl),
            Credentials = new WebCredentials(
                Settings.UserWithImpersonationAccess.Username,
                Settings.UserWithImpersonationAccess.Password
            ),
            ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, userCredentials.Username),
        };
        
        return service;
    }
    
    protected ExchangeService GetExchangeServiceUsingDelegatingAccess()
    {
        var service = new ExchangeService(GetWorkaroundTimeZone())
        {
            Url = new Uri(Settings.EwsServiceUrl),
            Credentials = new WebCredentials(
                Settings.UserWithDelegationAccess.Username,
                Settings.UserWithDelegationAccess.Password
            ),
        };
        
        return service;
    }
    
    private static TimeZoneInfo GetWorkaroundTimeZone()
    {
        return TimeZoneInfo.Utc;
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
            .AddUserSecrets<TestSettings>()
            .Build()
            .Get<TestSettings>();
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

    private void DeleteExistingAppointments(UserCredentials userCredentials)
    {
        var exchangeService = GetExchangeServiceUsingImpersonation(userCredentials);
        
        var calendar = CalendarFolder.Bind(exchangeService, WellKnownFolderName.Calendar, []);
        var calendarView = new CalendarView(DateTime.UtcNow.AddMonths(-1), DateTime.UtcNow.AddMonths(+1), int.MaxValue)
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

    /// <summary>
    /// Предоставляет доступ к календарю другому пользователю
    /// </summary>
    /// <param name="owner">Владелец календаря</param>
    /// <param name="service">Порльзователь, которму нужно предоставить доступ</param>
    /// <param name="folderPermissionLevel">Права</param>
    protected void GrantAccessToCalendar(
        UserCredentials owner, 
        UserCredentials service, 
        FolderPermissionLevel folderPermissionLevel = FolderPermissionLevel.Editor)
    {
        var exchangeService = GetExchangeServiceUsingDirectAccess(owner);
        
        var calendarFolder = Folder.Bind(
            exchangeService,
            WellKnownFolderName.Calendar,
            new PropertySet(BasePropertySet.IdOnly, FolderSchema.Permissions)
        );

        var permissions = calendarFolder.Permissions
            .FirstOrDefault(p => p.UserId.PrimarySmtpAddress == service.Username);
        if (permissions != null)
        {
            permissions.PermissionLevel = folderPermissionLevel;
        }
        else
        {
            calendarFolder.Permissions.Add(new FolderPermission(service.Username, folderPermissionLevel));
        }

        calendarFolder.Update();
    }
    
    /// <summary>
    /// Забирает доступ к календарю у другого пользователя
    /// </summary>
    /// <param name="owner">Владелец календаря</param>
    /// <param name="service">Порльзователь, у которого нужно забрать доступ</param>
    protected void RevokeAccessToCalendar(UserCredentials owner, UserCredentials service)
    {
        var exchangeService = GetExchangeServiceUsingDirectAccess(owner);
        
        var sentItemsFolder = Folder.Bind(
            exchangeService,
            WellKnownFolderName.Calendar, 
            new PropertySet(BasePropertySet.IdOnly, FolderSchema.Permissions)
        );
        
        var permissions = sentItemsFolder.Permissions
            .FirstOrDefault(p => p.UserId.PrimarySmtpAddress == service.Username);
        if (permissions != null)
        {
            sentItemsFolder.Permissions.Remove(permissions);
        }

        sentItemsFolder.Update();
    }

    [OneTimeTearDown]
    public void OneTimeTearDown()
    {
        // Clean up data
        DeleteExistingAppointments(Settings.User1);
        DeleteExistingAppointments(Settings.User2);
        DeleteExistingAppointments(Settings.User3);
        DeleteExistingAppointments(Settings.User4);
        DeleteExistingAppointments(Settings.User5);
        
        // Reset permissions
        RevokeAccessToCalendar(Settings.User1, Settings.UserWithDelegationAccess);
        RevokeAccessToCalendar(Settings.User2, Settings.UserWithDelegationAccess);
        RevokeAccessToCalendar(Settings.User3, Settings.UserWithDelegationAccess);
        RevokeAccessToCalendar(Settings.User4, Settings.UserWithDelegationAccess);
        RevokeAccessToCalendar(Settings.User5, Settings.UserWithDelegationAccess);

        // Так как тесты интеграционные, иногда что-то за чем-то не поспевает
        System.Threading.Tasks.Task.Delay(1000).Wait();
        
        // Dispose resources
        ServiceProvider.Dispose();
    }
}