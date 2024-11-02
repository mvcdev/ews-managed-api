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
    
    [Test]
    public void GetAttendeesAppointments()
    {
        // Arrange
        var exchangeService = GetExchangeService();
        
        var appointment = new Appointment(exchangeService)
        {
            Subject = Guid.NewGuid().ToString(),
            Start = DateTime.Now,
            End = DateTime.Now.AddHours(1),
            RequiredAttendees = { TestUsers.User2 },
            OptionalAttendees = { TestUsers.User3 },
        };
        appointment.Save(SendInvitationsMode.SendOnlyToAll);
        
        // Act
        var organizerCalendar = CalendarFolder.Bind(GetExchangeService(TestUsers.User1), WellKnownFolderName.Calendar, []);
        var requiredAttendeeCalendar = CalendarFolder.Bind(GetExchangeService(TestUsers.User2), WellKnownFolderName.Calendar, []);
        var optionalAttendeeCalendar = CalendarFolder.Bind(GetExchangeService(TestUsers.User3), WellKnownFolderName.Calendar, []);

        var calendarView = new CalendarView(DateTime.Now.Date, DateTime.Now.Date.AddDays(1), int.MaxValue)
        {
            PropertySet = new PropertySet(
                ItemSchema.Subject,
                AppointmentSchema.Organizer
            )
        };
        
        var organizerAppointments = organizerCalendar.FindAppointments(calendarView).ToArray();
        var requiredAttendeeAppointments = requiredAttendeeCalendar.FindAppointments(calendarView).ToArray();
        var optionalAttendeeAppointments = optionalAttendeeCalendar.FindAppointments(calendarView).ToArray();
        
        // Assert
        var organizerAppointment = (Appointment)Item.Bind(
            GetExchangeService(TestUsers.User1),
            organizerAppointments.First(a => a.Subject == appointment.Subject).Id,
            new PropertySet(
                ItemSchema.Subject,
                ItemSchema.ConversationId,
                AppointmentSchema.ICalUid,
                AppointmentSchema.Organizer
            ));
        organizerAppointment.Id.UniqueId.Should().Be(appointment.Id.UniqueId);
        organizerAppointment.Subject.Should().Be(appointment.Subject);
        organizerAppointment.Organizer.Name.Should().Be(TestUsers.User1.GetLogin());
        
        var requiredAttendeeAppointment = (Appointment)Item.Bind(
            GetExchangeService(TestUsers.User2),
            requiredAttendeeAppointments.First(a => a.Subject == appointment.Subject).Id,
            new PropertySet(
                ItemSchema.Subject,
                ItemSchema.ConversationId,
                AppointmentSchema.ICalUid,
                AppointmentSchema.Organizer
            ));
        requiredAttendeeAppointment.Id.UniqueId.Should().NotBe(appointment.Id.UniqueId, 
            "Мероприятия в календарях участников появляются с другим идентификатором");
        requiredAttendeeAppointment.Subject.Should().Be(appointment.Subject);
        requiredAttendeeAppointment.Organizer.Name.Should().Be(TestUsers.User1.GetLogin());
        requiredAttendeeAppointment.ICalUid.Should().Be(organizerAppointment.ICalUid);
        
        var optionalAttendeeAppointment = (Appointment)Item.Bind(
            GetExchangeService(TestUsers.User3),
            optionalAttendeeAppointments.First(a => a.Subject == appointment.Subject).Id,
            new PropertySet(
                ItemSchema.Subject,
                ItemSchema.ConversationId,
                AppointmentSchema.ICalUid,
                AppointmentSchema.Organizer
            ));
        optionalAttendeeAppointment.Id.UniqueId.Should().NotBe(appointment.Id.UniqueId, 
            "Мероприятия в календарях участников появляются с другим идентификатором");
        optionalAttendeeAppointment.Subject.Should().Be(appointment.Subject);
        optionalAttendeeAppointment.Organizer.Name.Should().Be(TestUsers.User1.GetLogin());
        optionalAttendeeAppointment.ICalUid.Should().Be(organizerAppointment.ICalUid);
    }
}