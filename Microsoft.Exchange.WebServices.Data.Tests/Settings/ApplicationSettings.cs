namespace Microsoft.Exchange.WebServices.Data.Tests.Settings;

public record ApplicationSettings
{
    public string EwsServiceUrl { get; init; } = string.Empty;
    public string Username { get; init; } = string.Empty;
    public string Password { get; init; } = string.Empty;
}