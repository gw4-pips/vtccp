namespace VtccpApp.Services;

/// <summary>
/// Provides the user-facing version from the compiled VtccpApp assembly.
/// Keeping this lookup in one place prevents the shell from drifting away
/// from the version declared in VtccpApp.csproj.
/// </summary>
internal static class AppVersionDisplay
{
    public static string Current =>
        Format(typeof(AppVersionDisplay).Assembly.GetName().Version);

    internal static string Format(Version? version)
    {
        if (version is null)
            throw new InvalidOperationException("The VtccpApp assembly has no version.");

        return version.Build >= 0
            ? $"v{version.Major}.{version.Minor}.{version.Build}"
            : $"v{version.Major}.{version.Minor}";
    }
}