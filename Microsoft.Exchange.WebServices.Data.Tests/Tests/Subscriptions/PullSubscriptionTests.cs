using System.Diagnostics;

namespace Microsoft.Exchange.WebServices.Data.Tests.Tests.Subscriptions;

public class PullSubscriptionTests : TestFixtureBase
{
    [Test]
    public void SubscribeToAppointments_Using_Impersonation()
    {
        // Arrange
        
        // Сначала начинаем создавать мероприятия
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
        
        // Потом создаем подписку
        var subscribedAppointments = new List<Appointment>();
        int batches = 0;
       
        // Затем подписываемся на уведомления
        var subscribeToAppointmentsTask = System.Threading.Tasks.Task.Run(() =>
        {
            var exchangeService = GetExchangeServiceUsingImpersonation(Settings.User1);
            var pullSubscription = exchangeService
                .SubscribeToPullNotifications([WellKnownFolderName.Calendar], 1, null, EventType.Created);
            while (true)
            {
                var eventsResult = pullSubscription.GetEvents();
                foreach (var @event in eventsResult.ItemEvents)
                {
                    var appointment = (Appointment)Item.Bind(exchangeService, @event.ItemId, new PropertySet());
                    subscribedAppointments.Add(appointment);
                }

                // Сделал public конструктор для PullSubscription,
                // который позволяет передать параметры существующей подписки
                // То есть SubscriptionId и Watermark можно хранить в БД
                // Таумайт подписки до 24ч
                // Но Exchange Client Services
                pullSubscription = new PullSubscription(
                    exchangeService,
                    pullSubscription.Id,
                    pullSubscription.Watermark,
                    pullSubscription.MoreEventsAvailable
                );

                batches++;

                if (subscribedAppointments.Count == 10)
                    break;
            }
            pullSubscription.Unsubscribe();
        });
        
        // Act
        System.Threading.Tasks.Task.WhenAll(
            createAppointmentsTask,
            subscribeToAppointmentsTask
        ).Wait(1000 * 20);
        
        // Assert
        batches.Should().BeGreaterThan(1, "Тест не очень корректный, так как подписка вернула все мероприятия за раз");
        subscribedAppointments.Count.Should().Be(createdAppointments.Count);
        subscribedAppointments.Should().Contain(sa => createdAppointments.Any(a => a.Id.UniqueId == sa.Id.UniqueId));
    }
    
    [Test]
    public void SubscribeToAppointments_Using_DelegatedAccess()
    {
        // Arrange
        GrantAccessToCalendar(Settings.User1, Settings.UserWithDelegationAccess);
        
        // Сначала начинаем создавать мероприятия
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
        
        // Потом создаем подписку
        var subscribedAppointments = new List<Appointment>();
        int batches = 0;
       
        // Затем подписываемся на уведомления
        var subscribeToAppointmentsTask = System.Threading.Tasks.Task.Run(() =>
        {
            var exchangeService = GetExchangeServiceUsingDelegatingAccess();
            var otherUserCalendar = new FolderId(WellKnownFolderName.Calendar, new Mailbox(Settings.User1.Username));
            
            var pullSubscription = exchangeService
                .SubscribeToPullNotifications([otherUserCalendar], 1, null, EventType.Created);
            while (true)
            {
                var eventsResult = pullSubscription.GetEvents();
                foreach (var @event in eventsResult.ItemEvents)
                {
                    var appointment = (Appointment)Item.Bind(exchangeService, @event.ItemId, new PropertySet());
                    subscribedAppointments.Add(appointment);
                }

                // Сделал public конструктор для PullSubscription,
                // который позволяет передать параметры существующей подписки
                // То есть SubscriptionId и Watermark можно хранить в БД
                // Таумайт подписки до 24ч
                // Но Exchange Client Services
                pullSubscription = new PullSubscription(
                    exchangeService,
                    pullSubscription.Id,
                    pullSubscription.Watermark,
                    pullSubscription.MoreEventsAvailable
                );

                batches++;

                if (subscribedAppointments.Count == 10)
                    break;
            }
            pullSubscription.Unsubscribe();
        });
        
        // Act
        System.Threading.Tasks.Task.WhenAll(
            createAppointmentsTask,
            subscribeToAppointmentsTask
        ).Wait(1000 * 20);
        
        // Assert
        batches.Should().BeGreaterThan(1, "Тест не очень корректный, так как подписка вернула все мероприятия за раз");
        subscribedAppointments.Count.Should().Be(createdAppointments.Count);
        subscribedAppointments.Should().Contain(sa => createdAppointments.Any(a => a.Id.UniqueId == sa.Id.UniqueId));
    }
    
    [Test]
    public void SubscribeToAppointments_Using_DelegatedAccess_ThroughSharedEmail()
    {
        // Arrange
        var organizer = Settings.User1;
        var sharedParticipant = Settings.User5;
        
        GrantAccessToCalendar(sharedParticipant, Settings.UserWithDelegationAccess);
        
        // Сначала начинаем создавать мероприятия
        var createdAppointments = new List<Appointment>();
        var createAppointmentsTask = System.Threading.Tasks.Task.Run(() =>
        {
            var exchangeService = GetExchangeServiceUsingImpersonation(organizer);
            for (var i = 1; i <= 10; i++)
            {
                System.Threading.Tasks.Task.Delay(500).Wait();
                
                var appointment = new Appointment(exchangeService)
                {
                    Subject = "Мероприятие " + i,
                    Body = "Сделать то, потом сделать сё",
                    Start = DateTime.UtcNow.AddHours(i),
                    End = DateTime.UtcNow.AddHours(i + 1),
                    Location = "Дома",
                    RequiredAttendees = { sharedParticipant.Username }
                };

                appointment.Save(SendInvitationsMode.SendOnlyToAll);
                
                createdAppointments.Add(appointment);
            }
        });
        
        // Потом создаем подписку
        var subscribedAppointments = new List<Appointment>();
        int batches = 0;
       
        // Затем подписываемся на уведомления
        var subscribeToAppointmentsTask = System.Threading.Tasks.Task.Run(() =>
        {
            var exchangeService = GetExchangeServiceUsingDelegatingAccess();
            var sharedCalendar = new FolderId(WellKnownFolderName.Calendar, sharedParticipant.Username);
            
            var pullSubscription = exchangeService
                .SubscribeToPullNotifications([sharedCalendar], 1, null, EventType.Created);
            while (true)
            {
                var eventsResult = pullSubscription.GetEvents();
                foreach (var @event in eventsResult.ItemEvents)
                {
                    var appointment = (Appointment)Item.Bind(exchangeService, @event.ItemId, new PropertySet(ItemSchema.Subject));
                    subscribedAppointments.Add(appointment);
                }

                // Сделал public конструктор для PullSubscription,
                // который позволяет передать параметры существующей подписки
                // То есть SubscriptionId и Watermark можно хранить в БД
                // Таумайт подписки до 24ч
                // Но Exchange Client Services
                pullSubscription = new PullSubscription(
                    exchangeService,
                    pullSubscription.Id,
                    pullSubscription.Watermark,
                    pullSubscription.MoreEventsAvailable
                );

                batches++;

                if (subscribedAppointments.Count == 10)
                    break;
            }
            pullSubscription.Unsubscribe();
        });
        
        // Act
        System.Threading.Tasks.Task.WhenAll(
            createAppointmentsTask,
            subscribeToAppointmentsTask
        ).Wait(1000 * 20);
        
        // Assert
        batches.Should().BeGreaterThan(1, "Тест не очень корректный, так как подписка вернула все мероприятия за раз");
        subscribedAppointments.Count.Should().Be(createdAppointments.Count);
        foreach (var subscribedAppointment in subscribedAppointments)
        {
            createdAppointments.Select(a => a.Subject).Should().Contain(id => id == subscribedAppointment.Subject);
        }
    }
    
    [Test]
    public void SubscribeToAppointments_Using_DelegatedAccess_ThroughSharedEmail_CRUD()
    {
        // Arrange
        var organizer = Settings.User1;
        var sharedParticipant = Settings.User5;
        
        GrantAccessToCalendar(sharedParticipant, Settings.UserWithDelegationAccess);
        
        // Сначала начинаем создавать мероприятия
        var createUpdateDeleteAppointmentsTask = System.Threading.Tasks.Task.Run(() =>
        {
            var exchangeService = GetExchangeServiceUsingImpersonation(organizer);
        
            // Создание
            System.Threading.Tasks.Task.Delay(3000).Wait();
            var appointment = new Appointment(exchangeService)
            {
                Subject = "Созданное мероприятие",
                Body = "Сделать то, потом сделать сё",
                Start = DateTime.UtcNow,
                End = DateTime.UtcNow.AddHours(1),
                Location = "Дома",
                RequiredAttendees = { sharedParticipant.Username }
            };
            appointment.Save(SendInvitationsMode.SendOnlyToAll);
            Debug.WriteLine("Мероприятие создано");
            
            // Редактирование
            System.Threading.Tasks.Task.Delay(3000).Wait();
            appointment.Subject = "Отредактированное мероприятие";
            appointment.Update(ConflictResolutionMode.AlwaysOverwrite);
            Debug.WriteLine("Мероприятие изменено");
            
            // Удаление
            System.Threading.Tasks.Task.Delay(3000).Wait();
            exchangeService.DeleteItems(
                [appointment.Id],
                DeleteMode.HardDelete,
                SendCancellationsMode.SendToNone,
                AffectedTaskOccurrence.AllOccurrences
            );
            Debug.WriteLine("Мероприятие удалено");
        });
        
        // Потом создаем подписку
        var events = new List<EventType>();
        var batches = 0;
       
        // Затем подписываемся на уведомления
        var subscribeToAppointmentsTask = System.Threading.Tasks.Task.Run(() =>
        {
            var exchangeService = GetExchangeServiceUsingDelegatingAccess();
            var sharedCalendar = new FolderId(WellKnownFolderName.Calendar, sharedParticipant.Username);
            
            var pullSubscription = exchangeService
                .SubscribeToPullNotifications([sharedCalendar], 1, null,
                    EventType.Created, EventType.Modified, EventType.Deleted);
            while (true)
            {
                var eventsResult = pullSubscription.GetEvents();
                foreach (var evt in eventsResult.ItemEvents)
                {
                    events.Add(evt.EventType);
                    Debug.WriteLine("Произошло событие " + evt.EventType);
                }

                // Сделал public конструктор для PullSubscription,
                // который позволяет передать параметры существующей подписки
                // То есть SubscriptionId и Watermark можно хранить в БД
                // Таумайт подписки до 24ч
                // Но Exchange Client Services
                pullSubscription = new PullSubscription(
                    exchangeService,
                    pullSubscription.Id,
                    pullSubscription.Watermark,
                    pullSubscription.MoreEventsAvailable
                );

                batches++;

                if (events.Count >= 3)
                    break;
            }
            pullSubscription.Unsubscribe();
        });
        
        // Act
        System.Threading.Tasks.Task.WhenAll(
            createUpdateDeleteAppointmentsTask,
            subscribeToAppointmentsTask
        ).Wait(1000 * 30);
        
        // Assert
        batches.Should().BeGreaterThan(1, "Тест не очень корректный, так как подписка вернула все мероприятия за раз");
        events.Should().ContainInOrder(
            EventType.Created,
            EventType.Modified,
            // todo: Удаление пока выглядит как модификация, надо будет разобраться с этим
            EventType.Modified
        );
        events.Count.Should().Be(3);
    }
}