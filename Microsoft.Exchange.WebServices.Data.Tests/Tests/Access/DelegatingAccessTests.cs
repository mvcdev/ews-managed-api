namespace Microsoft.Exchange.WebServices.Data.Tests.Tests.Access;

public class DelegatingAccessTests : TestFixtureBase
{
    [Test]
    public void CreateAppointment()
    {
        // Arrange
        var exchangeService = GetExchangeServiceUsingDelegatingAccess();
        
        var appointmentToCreate = new Appointment(exchangeService)
        {
            Subject = "Моё мероприятие",
            Body = "Сделать то, потом сделать сё",
            Start = DateTime.Now,
            End = DateTime.Now.AddHours(1),
            Location = "Дома"
        };
        
        var otherUserCalendar = new FolderId(WellKnownFolderName.Calendar, new Mailbox(TestUsers.User1));
        
        // Act
        appointmentToCreate.Save(otherUserCalendar, SendInvitationsMode.SendToNone);
        
        // Assert
        var appointment = (Appointment)Item.Bind(
            exchangeService,
            appointmentToCreate.Id,
            new PropertySet(
                ItemSchema.Subject,
                ItemSchema.Body,
                AppointmentSchema.Organizer,
                AppointmentSchema.Start,
                AppointmentSchema.End,
                AppointmentSchema.Location
            )
        );
        
        appointment.Organizer.Address.Should().Be(TestUsers.User1);
        appointment.Subject.Should().Be(appointmentToCreate.Subject);
        appointment.Body.Text.Should().Contain(appointmentToCreate.Body.Text);
        appointment.Start.Should().BeCloseTo(appointmentToCreate.Start, TimeSpan.FromSeconds(1));
        appointment.End.Should().BeCloseTo(appointmentToCreate.End, TimeSpan.FromSeconds(1));
        appointment.Location.Should().Be(appointmentToCreate.Location);
    }
    
    [Test]
    public void GetAppointmentList()
    {
        // Arrange
        var exchangeServiceWithImpersonationAccess = GetExchangeServiceUsingDelegatingAccess();
        
        var createdAppointment1 = new Appointment(exchangeServiceWithImpersonationAccess)
        {
            Subject = "Мероприятие 1",
            Start = DateTime.Now,
            End = DateTime.Now.AddHours(1),
            Location = "Дома"
        };
        createdAppointment1.Save(SendInvitationsMode.SendToNone);
        
        var createdAppointment2 = new Appointment(exchangeServiceWithImpersonationAccess)
        {
            Subject = "Мероприятие 2",
            Start = DateTime.Now.AddHours(2),
            End = DateTime.Now.AddHours(3),
            Location = "Дома"
        };
        createdAppointment2.Save(SendInvitationsMode.SendToNone);

        // Act
        var exchangeServiceWithDelegatingAccess = GetExchangeServiceUsingDelegatingAccess();

        var calendar = CalendarFolder.Bind(exchangeServiceWithDelegatingAccess, WellKnownFolderName.Calendar, []);

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