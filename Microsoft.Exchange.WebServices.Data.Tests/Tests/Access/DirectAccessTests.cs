namespace Microsoft.Exchange.WebServices.Data.Tests.Tests.Access;

public class DirectAccessTests : TestFixtureBase
{
    [Test]
    public void CreateAppointment()
    {
        // Arrange
        var exchangeService = GetExchangeServiceUsingDirectAccess(Settings.User5);
        var appointmentToCreate = new Appointment(exchangeService)
        {
            Subject = "Моё мероприятие",
            Body = "Сделать то, потом сделать сё",
            Start = DateTime.Now,
            End = DateTime.Now.AddHours(1),
            Location = "Дома"
        };

        // Act
        appointmentToCreate.Save(SendInvitationsMode.SendToNone);
        
        // Assert
        var appointment = (Appointment)Item.Bind(
            exchangeService,
            appointmentToCreate.Id,
            new PropertySet(
                ItemSchema.Subject,
                ItemSchema.Body,
                AppointmentSchema.Start,
                AppointmentSchema.End,
                AppointmentSchema.Location
            )
        );
        
        appointment.Subject.Should().Be(appointmentToCreate.Subject);
        appointment.Body.Text.Should().Contain(appointmentToCreate.Body.Text);
        appointment.Start.Should().BeCloseTo(appointmentToCreate.Start, TimeSpan.FromSeconds(1));
        appointment.End.Should().BeCloseTo(appointmentToCreate.End, TimeSpan.FromSeconds(1));
        appointment.Location.Should().Be(appointmentToCreate.Location);
    }
    
    [Test]
    public void GetAppointmentsList()
    {
        // Arrange
        var exchangeService = GetExchangeServiceUsingDirectAccess(Settings.User5);
        
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
        appointments.Should().NotBeNull();
    }
}