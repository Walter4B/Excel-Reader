using Microsoft.Extensions.Configuration;

namespace Excel_Reader
{
    internal class OutputController
    {
        private readonly IConfiguration configuration = GetConfiguration();

        internal void DisplayMessage(string messageKey)
        {
            Console.WriteLine(configuration[messageKey]);
        }

        internal static IConfiguration GetConfiguration()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("Excel Reader/appsettings.json");
            return builder.Build();
        }
    }
}
