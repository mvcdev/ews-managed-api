namespace Microsoft.Exchange.WebServices.Data.Tests.Tests;

public class GetAppointmentTests : TestFixtureBase
{
    [Test]
    public void GetAppointment()
    {
        // Arrange
        var exchangeService = GetExchangeServiceUsingImpersonation(Settings.User1);
        var appointmentToCreate = new Appointment(exchangeService)
        {
            Subject = "Моё мероприятие",
            Body = "Сделать то, потом сделать сё",
            Start = DateTime.UtcNow,
            End = DateTime.UtcNow.AddHours(1),
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
}