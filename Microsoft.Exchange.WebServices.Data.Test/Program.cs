using System.Collections;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Data.Test;
using Microsoft.Extensions.Configuration;
using Task = System.Threading.Tasks.Task;

// Start OWIN host 
ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;

// Настройка подключения
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

var variables = Environment.GetEnvironmentVariables();
foreach (DictionaryEntry variable in variables)
{
    Console.WriteLine(variable.Key + "=" + variable.Value);
}

var service = ExchangeServerExtensions
    .Configure(applicationSettings);

var appointmentsToDelete = service.GetAppointments(
    new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1),
    new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(1).AddDays(-1).AddTicks(-1),
    int.MaxValue);

// Создание уведомлений в фоновом потоке
if (false)
{
    service.DeleteAppointments(appointmentsToDelete.Select(a => a.Id).ToArray());

    _ = Task.Run(() =>
    {
        var i = 1;
        while (true)
        {
            service.CreateAppointment("Мероприятие " + i);
            Console.WriteLine("Created appointment " + i);
            i++;
            Task.Delay(3000).Wait();
        }
        // ReSharper disable once FunctionNeverReturns
    });
}

// Подписка на уведомления
if (false)
{
    var calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());
    var subscription = service.SubscribeToStreamingNotifications([calendar.Id],
        EventType.Created,
        EventType.Deleted,
        EventType.Modified,
        EventType.Moved,
        EventType.Copied,
        EventType.FreeBusyChanged
    );
    var connection = new StreamingSubscriptionConnection(service, 30);
    connection.AddSubscription(subscription);
    connection.OnNotificationEvent += (sender, args) =>
    {
        foreach (var @event in args.Events)
        {
            if (@event is not ItemEvent itemEvent) continue;

            var appointment = service.GetAppointment(itemEvent.ItemId, new PropertySet(
                ItemSchema.Subject,
                ItemSchema.LastModifiedTime,
                AppointmentSchema.Start,
                AppointmentSchema.End)
            );
            Console.WriteLine($"Subject: {appointment.Subject}; Start: {appointment.Start}; End: {appointment.End}");
        }
    };
    connection.OnDisconnect += (sender, args) => { };
    connection.Open();
}
// END OF: Подписка на уведомления


if (false)
{
    var pullSubscription = service.SubscribeToPullNotifications(
        [WellKnownFolderName.Calendar], 1, null, EventType.Created, EventType.Deleted, EventType.Modified,
        EventType.Moved);
    while (true)
    {
        Console.WriteLine(pullSubscription.Watermark);

        var eventsResult = pullSubscription.GetEvents();
        foreach (var @event in eventsResult.ItemEvents)
        {
            if (@event is not ItemEvent itemEvent) continue;

            var a = service.GetAppointment(itemEvent.ItemId, new PropertySet(
                ItemSchema.Subject,
                AppointmentSchema.Start,
                AppointmentSchema.End)
            );
            Console.WriteLine($"Received appointment: {a.Subject}");
        }

        // Сделал public конструктор для PullSubscription,
        // который позволяет передать параметры существующей подписки
        // То есть SubscriptionId и Watermark можно хранить в БД
        // Таумайт подписки до 24ч
        // Но Exchange Client Services
        pullSubscription = new PullSubscription(
            service,
            pullSubscription.Id,
            pullSubscription.Watermark,
            pullSubscription.MoreEventsAvailable
        );

        Task.Delay(3000).Wait();
    }

    pullSubscription.Unsubscribe();
}


// Создание мероприятия
var appointmentId = service.CreateAppointment("Моё мероприятие");

Console.WriteLine("Created appointment");
Console.ReadKey();
Task.Delay(5000).Wait();

// Получение мероприятия по Id
var appointment = service.GetAppointment(appointmentId, new PropertySet(
    ItemSchema.Subject,
    AppointmentSchema.Start,
    AppointmentSchema.End
));

// Редактирование мероприятия
service.UpdateAppointment(appointment);

Console.WriteLine("Updated appointment");
Task.Delay(5000).Wait();

// Получение списка мероприятий
// Initialize values for the start and end times, and the number of appointments to retrieve.
var startDate = new DateTime(appointment.Start.Year, appointment.Start.Month, 1);
var endDate = startDate.AddMonths(1).AddDays(-1).AddTicks(-1);
const int limit = int.MaxValue;

var appointments = service.GetAppointments(startDate, endDate, limit);
Console.WriteLine("\nThe first " + limit + " appointments on your calendar from " + startDate.Date.ToShortDateString() + 
                  " to " + endDate.Date.ToShortDateString() + " are: \n");

foreach (var a in appointments)
{
    Console.Write($"Subject: {a.Subject}; Start: {a.Start}; End: {a.End}");
    Console.WriteLine();
}

Console.WriteLine("Read appointment");
Task.Delay(5000).Wait();

service.DeleteAppointments(appointments.Select(a => a.Id).ToArray());

Console.WriteLine("Deleted appointment");
Task.Delay(5000).Wait();

Console.WriteLine($"Press any key to stop...");

Console.ReadKey();