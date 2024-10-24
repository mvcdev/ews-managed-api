using System.Net;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Data.Test;
using Microsoft.Extensions.Configuration;

// Start OWIN host 
ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;

// Настройка подключения
var applicationSettings = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json")
    .AddUserSecrets<ApplicationSettings>()
    .Build()
    .Get<ApplicationSettings>();

var service = ExchangeServerExtensions
    .Configure(applicationSettings ?? throw new InvalidOperationException());

// Создание мероприятия
var appointmentId = service.CreateAppointment();

// Получение мероприятия по Id
var appointment = service.GetAppointment(appointmentId, new PropertySet(
    ItemSchema.Subject,
    AppointmentSchema.Start,
    AppointmentSchema.End
));

// Редактирование мероприятия
service.UpdateAppointment(appointment);

// Получение списка мероприятий
// Initialize values for the start and end times, and the number of appointments to retrieve.
var startDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
var endDate = startDate.AddMonths(1).AddDays(-1).AddTicks(-1);
const int limit = int.MaxValue;

var appointments = service.GetAppointments(startDate, endDate, limit);
Console.WriteLine("\nThe first " + limit + " appointments on your calendar from " + startDate.Date.ToShortDateString() + 
                  " to " + endDate.Date.ToShortDateString() + " are: \n");

foreach (var a in appointments)
{
    Console.Write("Subject: " + a.Subject + "; ");
    Console.Write("Start: " + a.Start + "; ");
    Console.Write("End: " + a.End);
    Console.WriteLine();
}


service.DeleteAppointments(appointments.Select(a => a.Id).ToArray());

Console.WriteLine($"Press any key to stop...");

Console.ReadKey();