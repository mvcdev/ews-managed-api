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

var variables = Environment.GetEnvironmentVariables();
foreach (DictionaryEntry variable in variables)
{
    Console.WriteLine(variable.Key + "=" + variable.Value);
}

var service = ExchangeServerExtensions
    .Configure(applicationSettings ?? throw new InvalidOperationException());

// Создание мероприятия
var appointmentId = service.CreateAppointment();

Console.WriteLine("Created appointment");
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