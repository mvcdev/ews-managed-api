namespace Microsoft.Exchange.WebServices.Data.Tests.Settings;

public record ApplicationSettings
{
    public string EwsServiceUrl { get; init; } = string.Empty;
    public UserCredentials DirectAccess { get; init; } = null!;
    public UserCredentials Impersonation { get; init; } = null!;
    public UserCredentials DelegatingAccess { get; init; } = null!;
}

public record UserCredentials
{
    public string Username { get; init; } = string.Empty;
    public string Password { get; init; } = string.Empty;
}