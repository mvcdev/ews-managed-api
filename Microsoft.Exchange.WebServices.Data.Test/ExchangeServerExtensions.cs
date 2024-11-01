namespace Microsoft.Exchange.WebServices.Data.Test;

public static class ExchangeServerExtensions
{
    public static ExchangeService Configure(ApplicationSettings settings)
    {
        var service = new ExchangeService(GetWorkaroundTimeZone());
        
        service.Url = new Uri(settings.EwsServiceUrl);
            
        // todo у нас не будет пароля пользователя
        service.Credentials = new WebCredentials(settings.Username, settings.Password);

        service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "exuser1@airplan.local");
        
        Console.WriteLine($"Connecting to {settings.EwsServiceUrl} using credentials: {settings.Username} / {settings.Password}");
        
        return service;
    }

    private static TimeZoneInfo GetWorkaroundTimeZone()
    {
        // https://stackoverflow.com/questions/39467609/the-specified-time-zone-isnt-valid-using-ews-from-server
        return TimeZoneInfo.CreateCustomTimeZone(
            id: "Time zone to workaround a bug",
            baseUtcOffset: TimeZoneInfo.Local.BaseUtcOffset,
            displayName: "Time zone to workaround a bug", "Time zone to workaround a bug"
        );
    }
        
    public static ItemId CreateAppointment(this ExchangeService service, string subject)
    {
        var appointment = new Appointment(service);
        appointment.Subject = subject;
        appointment.Body = "Сделать то, потом сделать сё";
        appointment.Start = DateTime.UtcNow;
        appointment.End = appointment.Start.AddHours(1);
        appointment.Location = "Дома";
        appointment.ReminderDueBy = DateTime.UtcNow;

        appointment.RequiredAttendees.Add(new Attendee("exuser2@airplan.local"));
        appointment.OptionalAttendees.Add(new Attendee("exuser3@airplan.local"));
            
        // SendInvitationsMode
        //   * SendToNone - No meeting invitation is sent
        //   * SendOnlyToAll - Meeting invitations are sent to all attendees,
        //   * SendToAllAndSaveCopy - Meeting invitations are sent to all attendees and a copy of the invitation message is saved
        appointment.Save(SendInvitationsMode.SendToNone);
            
        return appointment.Id;
    }
        
    public static Appointment GetAppointment(this ExchangeService service, ItemId appointmentId, PropertySet propertiesToInclude)
    {
        return Item.Bind(service, appointmentId, propertiesToInclude) as Appointment;
    }
        
    public static void UpdateAppointment(this ExchangeService service, Appointment appointment)
    {
        appointment.Load(new PropertySet(
            ItemSchema.Subject,
            AppointmentSchema.Start,
            AppointmentSchema.End)
        );
        
        appointment.Subject += " 1"; 
        appointment.Start = appointment.Start.AddDays(1); 
        appointment.End = appointment.End.AddDays(1); 
            
        // ConflictResolutionMode
        //   * NeverOverwrite - Local property changes are discarded
        //   * AutoResolve - Local property changes are applied to the server unless the server-side copy is more recent than the local copy
        //   * AlwaysOverwrite - Local property changes overwrite server-side changes
        appointment.Update(ConflictResolutionMode.AlwaysOverwrite);
    }
        
    public static Appointment[] GetAppointments(this ExchangeService service, DateTime startDate, DateTime endDate, int limit)
    {
        var calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());
        var calendarView = new CalendarView(startDate, endDate, limit)
        {
            PropertySet = new PropertySet(ItemSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End)
        };

        return calendar.FindAppointments(calendarView).ToArray();
    }
        
    public static void DeleteAppointments(this ExchangeService service, params ItemId[] appointmentIds)
    {
        // DeleteMode
        //   * HardDelete - The item or folder will be permanently deleted
        //   * SoftDelete - The item or folder will be moved to the dumpster. Items and folders in the dumpster can be recovered.
        //   * MoveToDeletedItems - The item or folder will be moved to the mailbox' Deleted Items folder.
            
        // SendCancellationsMode
        //   * SendToNone - No meeting cancellation is sent.
        //   * SendOnlyToAll - Meeting cancellations are sent to all attendees.
        //   * SendToAllAndSaveCopy - Meeting cancellations are sent to all attendees and a copy of the cancellation message is saved in the organizer's Sent Items folder.
            
        // AffectedTaskOccurrence
        //   * AllOccurrences - All occurrences of the recurring task will be deleted.
        //   * SpecifiedOccurrenceOnly - Only the current occurrence of the recurring task will be deleted.

        service.DeleteItems(
            appointmentIds,
            DeleteMode.HardDelete,
            SendCancellationsMode.SendToNone,
            AffectedTaskOccurrence.AllOccurrences
        );
    }
}