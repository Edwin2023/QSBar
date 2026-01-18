namespace QSBar.Core.Security;

public static class SecurityManager
{
    static readonly HashSet<string> AllowedVersions = new HashSet<string> { "1.0", "2.0" };
    static readonly HashSet<string> AllowedUsers = new HashSet<string>();
    public static bool Enforce { get; set; } = false;

    public static bool IsVersionAllowed(string version)
    {
        return AllowedVersions.Contains(version);
    }

    public static void AllowVersion(string version)
    {
        AllowedVersions.Add(version);
    }

    public static bool IsAuthorized(string user, string version)
    {
        if (!IsVersionAllowed(version)) return false;
        if (AllowedUsers.Count == 0) return true;
        return AllowedUsers.Contains(user);
    }

    public static void ReloadFromConfig(SecurityData data)
    {
        AllowedVersions.Clear();
        foreach (var v in data.AllowedVersions) AllowedVersions.Add(v);
        AllowedUsers.Clear();
        foreach (var u in data.AllowedUsers) AllowedUsers.Add(u);
    }
}
