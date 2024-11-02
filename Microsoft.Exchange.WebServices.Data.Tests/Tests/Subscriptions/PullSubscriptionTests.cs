namespace Microsoft.Exchange.WebServices.Data.Tests.Tests.Subscriptions;

public class PullSubscriptionTests : TestFixtureBase
{
    [Test]
    public void SubscribeToAppointments()
    {
        // Arrange
        
        // Сначала начинаем создавать мероприятия
        var createdAppointments = new List<Appointment>();
        var createAppointmentsTask = System.Threading.Tasks.Task.Run(() =>
        {
            var exchangeService = GetExchangeService();
            for (var i = 1; i <= 10; i++)
            {
                System.Threading.Tasks.Task.Delay(500).Wait();
                
                var appointment = new Appointment(exchangeService)
                {
                    Subject = "Мероприятие " + i,
                    Body = "Сделать то, потом сделать сё",
                    Start = DateTime.Now.AddHours(i),
                    End = DateTime.Now.AddHours(i + 1),
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
            var exchangeService = GetExchangeService();
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
}