namespace Microsoft.Exchange.WebServices.Data.Tests.Tests.Subscriptions;

public class StreamingSubscriptionTests : TestFixtureBase
{
    [Test]
    public void SubscribeToAppointments_Using_Impersonation()
    {
        // Arrange
        
        // Сначала создаем подписку
        var subscriptionToAppointmentsTaskSource = new TaskCompletionSource();
        var subscribedAppointments = new List<Appointment>();
        var batches = 0;
        
        var exchangeService = GetExchangeServiceUsingImpersonation(Settings.User1);
        var calendar = CalendarFolder.Bind(exchangeService, WellKnownFolderName.Calendar, new PropertySet());
        var subscription = exchangeService.SubscribeToStreamingNotifications([calendar.Id], EventType.Created);
        var connection = new StreamingSubscriptionConnection(exchangeService, 30);
        connection.AddSubscription(subscription);
        connection.OnNotificationEvent += (sender, args) =>
        {
            foreach (var @event in args.Events)
            {
                if (@event is not ItemEvent itemEvent) continue;

                var appointment = (Appointment)Item.Bind(exchangeService, itemEvent.ItemId, new PropertySet());
                subscribedAppointments.Add(appointment);
            }

            batches++;
            
            if (subscribedAppointments.Count == 10)
                subscriptionToAppointmentsTaskSource.SetResult();
        };
        connection.OnDisconnect += (sender, args) => { };
        connection.Open();
        
        // Затем начинаем создавать мероприятия
        var createdAppointments = new List<Appointment>();
        var createAppointmentsTask = System.Threading.Tasks.Task.Run(() =>
        {
            var exchangeService = GetExchangeServiceUsingImpersonation(Settings.User1);
            for (var i = 1; i <= 10; i++)
            {
                System.Threading.Tasks.Task.Delay(500).Wait();
                
                var appointment = new Appointment(exchangeService)
                {
                    Subject = "Мероприятие " + i,
                    Body = "Сделать то, потом сделать сё",
                    Start = DateTime.UtcNow.AddHours(i),
                    End = DateTime.UtcNow.AddHours(i + 1),
                    Location = "Дома"
                };

                appointment.Save(SendInvitationsMode.SendToNone);
                
                createdAppointments.Add(appointment);
            }
            // ReSharper disable once FunctionNeverReturns
        });
        
        // Act
        System.Threading.Tasks.Task.WhenAll(
            createAppointmentsTask,
            subscriptionToAppointmentsTaskSource.Task
        ).Wait(1000 * 20);
        
        connection.Close();
        subscription.Unsubscribe();
        
        // Assert
        batches.Should().BeGreaterThan(1, "Тест не очень корректный, так как подписка вернула все мероприятия за раз");
        subscribedAppointments.Count.Should().Be(createdAppointments.Count);
        subscribedAppointments.Should().Contain(sa => createdAppointments.Any(a => a.Id.UniqueId == sa.Id.UniqueId));
    }
    
    [Test]
    public void SubscribeToAppointments_Using_DelegatingAccess()
    {
        // Arrange
        GrantAccessToCalendar(Settings.User1, Settings.UserWithDelegationAccess);
        
        // Сначала создаем подписку
        var subscriptionToAppointmentsTaskSource = new TaskCompletionSource();
        var subscribedAppointments = new List<Appointment>();
        var batches = 0;
        
        var exchangeService = GetExchangeServiceUsingDelegatingAccess();
        var otherUserCalendar = new FolderId(WellKnownFolderName.Calendar, new Mailbox(Settings.User1.Username));
        var calendar = CalendarFolder.Bind(exchangeService, otherUserCalendar, new PropertySet());
        var subscription = exchangeService.SubscribeToStreamingNotifications([calendar.Id], EventType.Created);
        var connection = new StreamingSubscriptionConnection(exchangeService, 30);
        connection.AddSubscription(subscription);
        connection.OnNotificationEvent += (sender, args) =>
        {
            foreach (var @event in args.Events)
            {
                if (@event is not ItemEvent itemEvent) continue;

                var appointment = (Appointment)Item.Bind(exchangeService, itemEvent.ItemId, new PropertySet());
                subscribedAppointments.Add(appointment);
            }

            batches++;
            
            if (subscribedAppointments.Count == 10)
                subscriptionToAppointmentsTaskSource.SetResult();
        };
        connection.OnDisconnect += (sender, args) => { };
        connection.Open();
        
        // Затем начинаем создавать мероприятия
        var createdAppointments = new List<Appointment>();
        var createAppointmentsTask = System.Threading.Tasks.Task.Run(() =>
        {
            var exchangeService = GetExchangeServiceUsingImpersonation(Settings.User1);
            for (var i = 1; i <= 10; i++)
            {
                System.Threading.Tasks.Task.Delay(500).Wait();
                
                var appointment = new Appointment(exchangeService)
                {
                    Subject = "Мероприятие " + i,
                    Body = "Сделать то, потом сделать сё",
                    Start = DateTime.UtcNow.AddHours(i),
                    End = DateTime.UtcNow.AddHours(i + 1),
                    Location = "Дома"
                };

                appointment.Save(SendInvitationsMode.SendToNone);
                
                createdAppointments.Add(appointment);
            }
            // ReSharper disable once FunctionNeverReturns
        });
        
        // Act
        System.Threading.Tasks.Task.WhenAll(
            createAppointmentsTask,
            subscriptionToAppointmentsTaskSource.Task
        ).Wait(1000 * 20);
        
        connection.Close();
        subscription.Unsubscribe();
        
        // Assert
        batches.Should().BeGreaterThan(1, "Тест не очень корректный, так как подписка вернула все мероприятия за раз");
        subscribedAppointments.Count.Should().Be(createdAppointments.Count);
        subscribedAppointments.Should().Contain(sa => createdAppointments.Any(a => a.Id.UniqueId == sa.Id.UniqueId));
    }
}