namespace Microsoft.Exchange.WebServices.Data.Tests.Tests;

public class GetAppointmentsListTests : TestFixtureBase
{
    /// <summary>
    /// Запрос списка мероприятий созданных одним пользователем
    /// </summary>
    [Test]
    public void GetAppointmentsList()
    {
        // Arrange
        var exchangeService = GetExchangeServiceUsingImpersonation(Settings.User1);
        
        var createdAppointment1 = new Appointment(exchangeService)
        {
            Subject = "Мероприятие 1",
            Start = DateTime.UtcNow,
            End = DateTime.UtcNow.AddHours(1),
            Location = "Дома"
        };
        createdAppointment1.Save(SendInvitationsMode.SendToNone);
        
        var createdAppointment2 = new Appointment(exchangeService)
        {
            Subject = "Мероприятие 2",
            Start = DateTime.UtcNow.AddHours(2),
            End = DateTime.UtcNow.AddHours(3),
            Location = "Дома"
        };
        createdAppointment2.Save(SendInvitationsMode.SendToNone);

        // Act
        var calendar = CalendarFolder.Bind(exchangeService, WellKnownFolderName.Calendar, []);

        var calendarView = new CalendarView(DateTime.UtcNow.Date, DateTime.UtcNow.Date.AddDays(1), int.MaxValue)
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
    
    /// <summary>
    /// Создание мероприятия с несколькими участниками,
    /// их получение через календари участников
    /// и сопостовление по одинаковому ICalUid
    /// </summary>
    [Test]
    public void GetAttendeesAppointments()
    {
        // Arrange
        var exchangeService = GetExchangeServiceUsingImpersonation(Settings.User1);
        
        var appointment = new Appointment(exchangeService)
        {
            Subject = Guid.NewGuid().ToString(),
            Start = DateTime.UtcNow,
            End = DateTime.UtcNow.AddHours(1),
            RequiredAttendees = { Settings.User2.Username },
            OptionalAttendees = { Settings.User3.Username },
        };
        appointment.Save(SendInvitationsMode.SendOnlyToAll);

        System.Threading.Tasks.Task.Delay(1000).Wait();
        
        // Act
        var calendarView = new CalendarView(DateTime.UtcNow.Date, DateTime.UtcNow.Date.AddDays(1), int.MaxValue)
        {
            PropertySet = new PropertySet(
                ItemSchema.Subject,
                AppointmentSchema.Organizer
            )
        };
        
        var organizerCalendar = CalendarFolder.Bind(GetExchangeServiceUsingImpersonation(Settings.User1), WellKnownFolderName.Calendar, []);
        var organizerAppointments = organizerCalendar.FindAppointments(calendarView).ToArray();
        
        var requiredAttendeeCalendar = CalendarFolder.Bind(GetExchangeServiceUsingImpersonation(Settings.User2), WellKnownFolderName.Calendar, []);
        var requiredAttendeeAppointments = requiredAttendeeCalendar.FindAppointments(calendarView).ToArray();
        
        var optionalAttendeeCalendar = CalendarFolder.Bind(GetExchangeServiceUsingImpersonation(Settings.User3), WellKnownFolderName.Calendar, []);
        var optionalAttendeeAppointments = optionalAttendeeCalendar.FindAppointments(calendarView).ToArray();
        
        // Assert
        var organizerAppointment = (Appointment)Item.Bind(
            GetExchangeServiceUsingImpersonation(Settings.User1),
            organizerAppointments.First(a => a.Subject == appointment.Subject).Id,
            new PropertySet(
                ItemSchema.Subject,
                ItemSchema.ConversationId,
                AppointmentSchema.ICalUid,
                AppointmentSchema.Organizer
            ));
        organizerAppointment.Id.UniqueId.Should().Be(appointment.Id.UniqueId);
        organizerAppointment.Subject.Should().Be(appointment.Subject);
        organizerAppointment.Organizer.Name.Should().Be(Settings.User1.Username.GetLogin());
        
        var requiredAttendeeAppointment = (Appointment)Item.Bind(
            GetExchangeServiceUsingImpersonation(Settings.User2),
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
        requiredAttendeeAppointment.Organizer.Name.Should().Be(Settings.User1.Username.GetLogin());
        requiredAttendeeAppointment.ICalUid.Should().Be(organizerAppointment.ICalUid);
        
        var optionalAttendeeAppointment = (Appointment)Item.Bind(
            GetExchangeServiceUsingImpersonation(Settings.User3),
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
        optionalAttendeeAppointment.Organizer.Name.Should().Be(Settings.User1.Username.GetLogin());
        optionalAttendeeAppointment.ICalUid.Should().Be(organizerAppointment.ICalUid);
    }
    
    /// <summary>
    /// Добавление пользователя опциональным или обязательным участников в мероприятие
    /// и отслеживание изменений в мероприятии в календаре этого участника
    /// </summary>
    [Test]
    public void GetAppointmentsThroughSharedCalendar()
    {
        // Arrange
        var user1Appointment = new Appointment(GetExchangeServiceUsingImpersonation(Settings.User1))
        {
            Subject = Guid.NewGuid().ToString(),
            Start = DateTime.UtcNow,
            End = DateTime.UtcNow.AddHours(1),
            OptionalAttendees = { Settings.User5.Username }
        };
        user1Appointment.Save(SendInvitationsMode.SendOnlyToAll);
        user1Appointment.Load(new PropertySet(ItemSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.ICalUid));
        
        var user2Appointment = new Appointment(GetExchangeServiceUsingImpersonation(Settings.User1))
        {
            Subject = Guid.NewGuid().ToString(),
            Start = DateTime.UtcNow,
            End = DateTime.UtcNow.AddHours(1),
            RequiredAttendees = { Settings.User5.Username },
        };
        user2Appointment.Save(SendInvitationsMode.SendOnlyToAll);
        user2Appointment.Load(new PropertySet(ItemSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.ICalUid));
        
        // todo Иногда мероприятия не успевают появиться в списке. Надо будет написать какую-то wait-обертку для получения списка
        System.Threading.Tasks.Task.Delay(1000).Wait();
        
        // Act
        var sharedCalendar = CalendarFolder.Bind(GetExchangeServiceUsingImpersonation(Settings.User5), WellKnownFolderName.Calendar, []);
        var calendarView = new CalendarView(DateTime.UtcNow.Date, DateTime.UtcNow.Date.AddDays(1), int.MaxValue)
        {
            PropertySet = new PropertySet(
                ItemSchema.Subject,
                AppointmentSchema.Organizer
            )
        };
        var appointments = sharedCalendar.FindAppointments(calendarView).ToArray();
        
        // Assert
        var optionalAttendeeAppointment = (Appointment)Item.Bind(
            GetExchangeServiceUsingImpersonation(Settings.User5),
            appointments.First(a => a.Subject == user1Appointment.Subject).Id,
            new PropertySet(
                ItemSchema.Subject,
                ItemSchema.ConversationId,
                AppointmentSchema.ICalUid,
                AppointmentSchema.Organizer
            ));
        optionalAttendeeAppointment.Subject.Should().Be(user1Appointment.Subject);
        optionalAttendeeAppointment.Organizer.Name.Should().Be(Settings.User1.Username.GetLogin());
        optionalAttendeeAppointment.ICalUid.Should().Be(user1Appointment.ICalUid);
        
        var requiredAttendeeAppointment = (Appointment)Item.Bind(
            GetExchangeServiceUsingImpersonation(Settings.User5),
            appointments.First(a => a.Subject == user2Appointment.Subject).Id,
            new PropertySet(
                ItemSchema.Subject,
                ItemSchema.ConversationId,
                AppointmentSchema.ICalUid,
                AppointmentSchema.Organizer
            ));
        requiredAttendeeAppointment.Subject.Should().Be(user2Appointment.Subject);
        requiredAttendeeAppointment.Organizer.Name.Should().Be(Settings.User1.Username.GetLogin());
        requiredAttendeeAppointment.ICalUid.Should().Be(user2Appointment.ICalUid);
    }
    
    /// <summary>
    /// Запрос списка мероприятий созданных одним пользователем
    /// </summary>
    [Test]
    public void GetAppointmentsListWithLastModifiedTime()
    {
        // Arrange
        var exchangeService = GetExchangeServiceUsingImpersonation(Settings.User1);

        var createdOn = DateTime.UtcNow;
        
        var createdAppointment1 = new Appointment(exchangeService)
        {
            Subject = "Мероприятие 1",
            Start = DateTime.UtcNow,
            End = DateTime.UtcNow.AddHours(1),
            Location = "Дома"
        };
        createdAppointment1.Save(SendInvitationsMode.SendToNone);
        
        var createdAppointment2 = new Appointment(exchangeService)
        {
            Subject = "Мероприятие 2",
            Start = DateTime.UtcNow.AddHours(2),
            End = DateTime.UtcNow.AddHours(3),
            Location = "Дома"
        };
        
        System.Threading.Tasks.Task.Delay(1000).Wait();
        
        createdAppointment2.Save(SendInvitationsMode.SendToNone);

        // Act
        var calendar = CalendarFolder.Bind(exchangeService, WellKnownFolderName.Calendar, []);

        var calendarView = new CalendarView(DateTime.UtcNow.Date, DateTime.UtcNow.Date.AddDays(1), int.MaxValue)
        {
            PropertySet = new PropertySet(
                ItemSchema.Subject,
                ItemSchema.LastModifiedTime,
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
        appointment1.LastModifiedTime.Should().BeAfter(createdOn);
        
        var appointment2 = appointments.First(a => a.Id.UniqueId == createdAppointment2.Id.UniqueId);
        appointment2.Subject.Should().Be(createdAppointment2.Subject);
        appointment2.Start.Should().BeCloseTo(createdAppointment2.Start, TimeSpan.FromSeconds(1));
        appointment2.End.Should().BeCloseTo(createdAppointment2.End, TimeSpan.FromSeconds(1));
        appointment2.Location.Should().Be(createdAppointment2.Location);
        appointment2.LastModifiedTime.Should().BeAfter(createdOn);
    }
}