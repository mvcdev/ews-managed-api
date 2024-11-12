namespace Microsoft.Exchange.WebServices.Data.Tests.Tests;

public class CreateAppointmentsTests : TestFixtureBase
{
    [Test]
    public void CreateAppointment()
    {
        // Arrange
        var exchangeService = GetExchangeServiceUsingImpersonation(Settings.User1);
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
    public void CreateAppointmentWithAttendees()
    {
        // Arrange
        var exchangeService = GetExchangeServiceUsingImpersonation(Settings.User1);
        var appointmentToCreate = new Appointment(exchangeService)
        {
            Subject = "Моё мероприятие",
            Start = DateTime.Now,
            End = DateTime.Now.AddHours(1),
            RequiredAttendees = { Settings.User2.Username },
            OptionalAttendees = { Settings.User3.Username },
        };

        // Act
        appointmentToCreate.Save(SendInvitationsMode.SendToNone);
        
        // Assert
        var appointment = (Appointment)Item.Bind(
            exchangeService,
            appointmentToCreate.Id,
            new PropertySet(
                AppointmentSchema.Organizer,
                AppointmentSchema.OptionalAttendees,
                AppointmentSchema.RequiredAttendees
            )
        );
        
        appointment.Organizer.Address.Should().Be(Settings.User1.Username);
        appointment.RequiredAttendees.Should().ContainSingle(a => a.Address == Settings.User2.Username);
        appointment.OptionalAttendees.Should().ContainSingle(a => a.Address == Settings.User3.Username);
    }
}