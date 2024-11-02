namespace Microsoft.Exchange.WebServices.Data.Tests.Tests;

public class GetAppointmentsListTests : TestFixtureBase
{
    [Test]
    public void GetAppointmentsList()
    {
        // Arrange
        var exchangeService = GetExchangeService();
        
        var createdAppointment1 = new Appointment(exchangeService)
        {
            Subject = "Мероприятие 1",
            Start = DateTime.Now,
            End = DateTime.Now.AddHours(1),
            Location = "Дома"
        };
        createdAppointment1.Save(SendInvitationsMode.SendToNone);
        
        var createdAppointment2 = new Appointment(exchangeService)
        {
            Subject = "Мероприятие 2",
            Start = DateTime.Now.AddHours(2),
            End = DateTime.Now.AddHours(3),
            Location = "Дома"
        };
        createdAppointment2.Save(SendInvitationsMode.SendToNone);

        // Act
        var calendar = CalendarFolder.Bind(exchangeService, WellKnownFolderName.Calendar, []);

        var calendarView = new CalendarView(DateTime.Now.Date, DateTime.Now.Date.AddDays(1), int.MaxValue)
        {
            PropertySet = new PropertySet(
                ItemSchema.Subject,
                AppointmentSchema.Start,
                AppointmentSchema.End,
                AppointmentSchema.Location
            )
        };

        var appointments = calendar.FindAppointments(calendarView).ToArray();
        
        // Assert
        var appointment1 = appointments.First(a => a.Id.UniqueId == createdAppointment1.Id.UniqueId);
        appointment1.Subject.Should().Be(createdAppointment1.Subject);
        appointment1.Start.Should().BeCloseTo(createdAppointment1.Start, TimeSpan.FromSeconds(1));
        appointment1.End.Should().BeCloseTo(createdAppointment1.End, TimeSpan.FromSeconds(1));
        appointment1.Location.Should().Be(createdAppointment1.Location);
        
        var appointment2 = appointments.First(a => a.Id.UniqueId == createdAppointment2.Id.UniqueId);
        appointment2.Subject.Should().Be(createdAppointment2.Subject);
        appointment2.Start.Should().BeCloseTo(createdAppointment2.Start, TimeSpan.FromSeconds(1));
        appointment2.End.Should().BeCloseTo(createdAppointment2.End, TimeSpan.FromSeconds(1));
        appointment2.Location.Should().Be(createdAppointment2.Location);
    }
}