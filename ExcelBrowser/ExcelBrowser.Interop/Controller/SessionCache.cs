using System.Collections.Generic;
using System.Linq;
using ExcelBrowser.Interop;
using ExcelBrowser.Model;

namespace ExcelBrowser.Controller {

    internal class SessionCache {

        public SessionCache(Session session) {
            Requires.NotNull(session, nameof(session));
            this.session = session;

        }

        private readonly Session session;
        private SessionToken previousSession;

        public SessionToken Current {
            get {
                IEnumerable<AppToken> reachableApps = session.Apps
                    .Select(TokenFactory.App)
                    .OrderBy(at => at.Id);

                var unreachableApps = session.UnreachableProcessIds
                    .Select(id => TokenFactory.InvisibleApp(new AppId(id)))
                    .OrderBy(a => a.Id);

                var apps = ReplaceInvisibleApps(
                    reachableApps.Concat(unreachableApps));

                var primary = session.Primary;
                var primaryId = (primary != null) ? IdFactory.App(primary) : null;

                var result = new SessionToken(
                    id: new SessionId(session.SessionId),
                    apps: apps,
                    primaryAppId: primaryId);

                this.previousSession = result;
                return result;
            }
        }

        //Replaces any busy apps that were previously cached with their previous version
        private IEnumerable<AppToken> ReplaceInvisibleApps(IEnumerable<AppToken> apps) {
            if (this.previousSession == null) {
                //Don't do anything special before first snapshot is saved
                foreach (var a in apps) yield return a;
            }
            else {
                var previousApps = this.previousSession.Apps.ToArray();

                foreach (var a in apps) {
                    if (a.IsVisible) { //If visible, return input
                        yield return a;
                    }
                    else { //Return previous match, but default to input
                        var prev = previousApps.SingleOrDefault(p => Equals(p.Id, a.Id));
                        if (prev != null) { //Mark and return previous match
                            prev = prev.ShallowCopy;
                            prev.IsVisible = false;
                            yield return prev;
                        }
                        else { //If no previous match, return input
                            yield return a;
                        }
                    }
                }
            }
        }

    }
}
