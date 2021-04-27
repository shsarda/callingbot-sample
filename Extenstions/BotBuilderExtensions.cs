namespace Microsoft.Extensions.DependencyInjection
{
    using CallingMeetingBot.Bot;
    using System;
    public static class BotBuilderExtensions
    {
        public static IServiceCollection AddBot(this IServiceCollection services)
            => services.AddBot(_ => { });

        public static IServiceCollection AddBot(this IServiceCollection services, Action<BotOptions> botOptionsAction)
        {
            var options = new BotOptions();
            botOptionsAction(options);
            services.AddSingleton(options);

            return services.AddSingleton<CallingAndMeetingBot>();
        }
    }
}
