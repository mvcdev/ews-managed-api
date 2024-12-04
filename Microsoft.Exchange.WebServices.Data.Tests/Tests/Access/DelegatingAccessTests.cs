namespace Microsoft.Exchange.WebServices.Data.Tests.Tests.Access;

public class DelegatingAccessTests : TestFixtureBase
{
    [Test]
    public void ShouldThrowIfNoAccess()
    {
        // Arrange
        RevokeAccessToCalendar(Settings.User5, Settings.UserWithDelegationAccess);
            
        var exchangeService = GetExchangeServiceUsingDelegatingAccess();
        
        var otherUserCalendar = new FolderId(WellKnownFolderName.Calendar, new Mailbox(Settings.User5.Username));
        
        // Act
        var getAppointments = () =>
        {
            var calendar = CalendarFolder.Bind(exchangeService, otherUserCalendar, []);
            
            var calendarView = new CalendarView(DateTime.UtcNow.Date, DateTime.UtcNow.Date.AddDays(1), int.MaxValue)
            {
                PropertySet = new PropertySet(
                    ItemSchema.Subject,
                    AppointmentSchema.Start,
                    AppointmentSchema.End,
                    AppointmentSchema.Location
                )
            };
            
            calendar.FindAppointments(calendarView);
        };
        
        // Assert
        getAppointments.Should().Throw<ServiceResponseException>();
    }
    
    [Test]
    public void ShouldNotThrowIfAccessGranted()
    {
        // Arrange
        GrantAccessToCalendar(Settings.User5, Settings.UserWithDelegationAccess);
        
        var exchangeService = GetExchangeServiceUsingDelegatingAccess();
        
        // Act
        var getAppointments = () =>
        {
            var otherUserCalendar = new FolderId(WellKnownFolderName.Calendar, new Mailbox(Settings.User5.Username));
            
            var calendar = CalendarFolder.Bind(exchangeService, otherUserCalendar, []);
            
            var calendarView = new CalendarView(DateTime.UtcNow.Date, DateTime.UtcNow.Date.AddDays(1), int.MaxValue)
            {
                PropertySet = new PropertySet(
                    ItemSchema.Subject,
                    AppointmentSchema.Start,
                    AppointmentSchema.End,
                    AppointmentSchema.Location
                )
            };
            
            calendar.FindAppointments(calendarView);
        };
        
        // Assert
        getAppointments.Should().NotThrow();
    }
    
    [Test]
    public void CreateAppointment()
    {
        // Arrange
        GrantAccessToCalendar(Settings.User5, Settings.UserWithDelegationAccess);
        
        var exchangeService = GetExchangeServiceUsingDelegatingAccess();
        
        var appointmentToCreate = new Appointment(exchangeService)
        {
            Subject = "Моё мероприятие",
            Body = "Сделать то, потом сделать сё",
            Start = DateTime.UtcNow,
            End = DateTime.UtcNow.AddHours(1),
            Location = "Дома"
        };
        
        var otherUserCalendar = new FolderId(WellKnownFolderName.Calendar, new Mailbox(Settings.User5.Username));
        
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
        
        appointment.Organizer.Address.Should().Be(Settings.User5.Username);
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
            Start = DateTime.UtcNow,
            End = DateTime.UtcNow.AddHours(1),
            Location = "Дома"
        };
        createdAppointment1.Save(SendInvitationsMode.SendToNone);
        
        var createdAppointment2 = new Appointment(exchangeServiceWithImpersonationAccess)
        {
            Subject = "Мероприятие 2",
            Start = DateTime.UtcNow.AddHours(2),
            End = DateTime.UtcNow.AddHours(3),
            Location = "Дома"
        };
        createdAppointment2.Save(SendInvitationsMode.SendToNone);

        // Act
        var exchangeServiceWithDelegatingAccess = GetExchangeServiceUsingDelegatingAccess();

        var calendar = CalendarFolder.Bind(exchangeServiceWithDelegatingAccess, WellKnownFolderName.Calendar, []);

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
    
    [Test]
    public void EditAppointmentCreatedByYourself_UsingAuthorPermissionLevel()
    {
        // Arrange
        var author = Settings.User5;
        
        GrantAccessToCalendar(author, Settings.UserWithDelegationAccess, FolderPermissionLevel.Author);
        
        var exchangeService = GetExchangeServiceUsingDelegatingAccess();
        
        var createdAppointment = new Appointment(exchangeService)
        {
            Subject = "Моё мероприятие",
            Body = "Сделать то, потом сделать сё",
            Start = DateTime.UtcNow,
            End = DateTime.UtcNow.AddHours(1),
            Location = "Дома"
        };
        
        var otherUserCalendar = new FolderId(WellKnownFolderName.Calendar, new Mailbox(Settings.User5.Username));
        
        createdAppointment.Save(otherUserCalendar, SendInvitationsMode.SendToNone);
        
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
    
    [Test]
    public void EditAppointmentCreatedByOther_UsingAuthorPermissionLevel()
    {
        // Arrange
        var author = Settings.User5;
        var authorExchangeService = GetExchangeServiceUsingDirectAccess(author);
        
        var createdAppointment = new Appointment(authorExchangeService)
        {
            Subject = "Моё мероприятие",
            Body = "Сделать то, потом сделать сё",
            Start = DateTime.UtcNow,
            End = DateTime.UtcNow.AddHours(1),
            Location = "Дома"
        };
        
        var otherUserCalendar = new FolderId(WellKnownFolderName.Calendar, new Mailbox(Settings.User5.Username));
        
        createdAppointment.Save(otherUserCalendar, SendInvitationsMode.SendToNone);
        
        // Act
        GrantAccessToCalendar(author, Settings.UserWithDelegationAccess, FolderPermissionLevel.Author);
        
        var delegatedAccessWithAuthorPermissions = GetExchangeServiceUsingDelegatingAccess();
        
        var appointmentToUpdate = (Appointment)Item.Bind(
            delegatedAccessWithAuthorPermissions,
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

        var updateAppointment = () => appointmentToUpdate.Update(ConflictResolutionMode.AlwaysOverwrite);
        
        // Assert
        updateAppointment.Should().Throw<ServiceResponseException>()
            .WithMessage(
                "Access is denied. Check credentials and try again., Cannot save changes made to an item to store.");
    }
}