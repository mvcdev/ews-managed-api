namespace Microsoft.Exchange.WebServices.Data.Tests.Tests;

public class UpdateAppointmentsTests : TestFixtureBase
{
    [Test]
    public void UpdateAppointment()
    {
        // Arrange
        var exchangeService = GetExchangeServiceUsingImpersonation(Settings.User1);
        var createdAppointment = new Appointment(exchangeService)
        {
            Subject = "Моё мероприятие",
            Body = "Сделать то, потом сделать сё",
            Start = DateTime.Now,
            End = DateTime.Now.AddHours(1),
            Location = "Дома"
        };
        
        createdAppointment.Save(SendInvitationsMode.SendToNone);
        
        // Act
        var appointmentToUpdate = (Appointment)Item.Bind(
            exchangeService,
            createdAppointment.Id,
            new PropertySet(
                ItemSchema.Subject,
                ItemSchema.Body,
                AppointmentSchema.Start,
                AppointmentSchema.End,
                AppointmentSchema.Location
            )
        );
        
        appointmentToUpdate.Subject = "Новое название моего мероприятия"; 
        appointmentToUpdate.Body = "Новое описание моего мероприятия"; 
        appointmentToUpdate.Start = appointmentToUpdate.Start.AddDays(1); 
        appointmentToUpdate.End = appointmentToUpdate.End.AddDays(1); 
        appointmentToUpdate.Location = "В офисе";
        
        appointmentToUpdate.Update(ConflictResolutionMode.AlwaysOverwrite);
        
        // Assert
        var updatedAppointment = (Appointment)Item.Bind(
            exchangeService,
            appointmentToUpdate.Id,
            new PropertySet(
                ItemSchema.Subject,
                ItemSchema.Body,
                AppointmentSchema.Start,
                AppointmentSchema.End,
                AppointmentSchema.Location
            )
        );
        
        updatedAppointment.Subject.Should().Be(appointmentToUpdate.Subject);
        updatedAppointment.Body.Text.Should().Contain(appointmentToUpdate.Body.Text);
        updatedAppointment.Start.Should().BeCloseTo(appointmentToUpdate.Start, TimeSpan.FromSeconds(1));
        updatedAppointment.End.Should().BeCloseTo(appointmentToUpdate.End, TimeSpan.FromSeconds(1));
        updatedAppointment.Location.Should().Be(appointmentToUpdate.Location);
    }
}