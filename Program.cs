using appsvc_fnc_dev_CreateUser_dotnet;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices(services =>
    {
        services.AddLogging();
        services.AddSingleton<Auth>();
        services.AddScoped<QueueUserInfo>();
        services.AddScoped<SendEmailQueueTrigger>();
    })
    .Build();

host.Run();