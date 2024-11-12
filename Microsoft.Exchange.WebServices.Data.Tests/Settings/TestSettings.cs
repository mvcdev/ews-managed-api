namespace Microsoft.Exchange.WebServices.Data.Tests.Settings;

public record TestSettings
{
    public string EwsServiceUrl { get; init; } = string.Empty;
    
    public UserCredentials UserWithImpersonationAccess { get; init; } = null!;
    public UserCredentials UserWithDelegationAccess { get; init; } = null!;
    public UserCredentials User1 { get; init; } = null!;
    public UserCredentials User2 { get; init; } = null!;
    public UserCredentials User3 { get; init; } = null!;
    public UserCredentials User4 { get; init; } = null!;
    public UserCredentials User5 { get; init; } = null!;
}

public record UserCredentials
{
    public string Username { get; init; } = string.Empty;
    public string Password { get; init; } = string.Empty;
}