namespace Microsoft.Exchange.WebServices.Data.Tests.Core;

public static class UserExtensions
{
    /// <summary>
    /// Возвращает логин пользователя по его почте 
    /// </summary>
    public static string GetLogin(this string? smtpAddress)
    {
        if (string.IsNullOrWhiteSpace(smtpAddress))
            return string.Empty;
        
        var parts = smtpAddress.Split('@');
        if (parts.Length == 2)
        {
            return parts[0];
        }

        return string.Empty;
    }
}