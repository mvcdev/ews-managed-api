namespace Microsoft.Exchange.WebServices.Data.Tests.Tests;

public class DeleteAppointmentTests : TestFixtureBase
{
    [Test]
    public void DeleteAppointment()
    {
        // Arrange
        var exchangeService = GetExchangeServiceUsingImpersonation();
        var appointment = new Appointment(exchangeService)
        {
            Subject = "Моё мероприятие",
            Body = "Сделать то, потом сделать сё",
            Start = DateTime.Now,
            End = DateTime.Now.AddHours(1),
            Location = "Дома"
        };
        appointment.Save(SendInvitationsMode.SendToNone);

        // Act
        exchangeService.DeleteItems(
            [appointment.Id],
            DeleteMode.HardDelete,
            SendCancellationsMode.SendToNone,
            AffectedTaskOccurrence.AllOccurrences
        );

        // Assert
        var calendar = CalendarFolder.Bind(exchangeService, WellKnownFolderName.Calendar, []);
        var calendarView = new CalendarView(DateTime.Now.Date, DateTime.Now.Date.AddDays(1), int.MaxValue)
        {
            PropertySet = new PropertySet(ItemSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End)
        };

        var appointments = calendar.FindAppointments(calendarView).ToArray();
        appointments.Should().NotContain(a => a.Id.UniqueId == appointment.Id.UniqueId);
    }
}