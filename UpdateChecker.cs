using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Reflection;
using Octokit;

namespace PPTXcreator
{
    static class UpdateChecker
    {
        public static async Task CheckReleases()
        {
            // Get all releases
            GitHubClient client = new GitHubClient(new ProductHeaderValue("PPTXcreator"));
            IReadOnlyList<Release> releases = await client.Repository.Release.GetAll("Sionkerk-Techniek", "PPTXcreator");

            // Get latest version number
            Version latestGithubVersion;
            try
            {
                latestGithubVersion = new Version(releases[0].TagName.Trim('v'));
            }
            catch (ArgumentOutOfRangeException) // No releases have been made yet
            {
                return;
            }

            // Get current version, without revision number since that is not used in the Github tags
            Version localVersion = Assembly.GetExecutingAssembly().GetName().Version;
            Version localVersionWithoutRevision = new Version(localVersion.Major, localVersion.Minor, localVersion.Build);

            if (localVersionWithoutRevision.CompareTo(latestGithubVersion) < 0)
            {
                Dialogs.UpdateAvailable(latestGithubVersion.ToString(3));
            }
        }
    }
}
