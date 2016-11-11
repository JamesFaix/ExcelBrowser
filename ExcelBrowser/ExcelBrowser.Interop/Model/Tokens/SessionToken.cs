using System;
using System.Collections.Immutable;
using System.Collections.Generic;
using System.Linq;
using ExcelBrowser.Interop;
using System.Diagnostics;

namespace ExcelBrowser.Model {

    public class SessionToken : Token<SessionId>, IEquatable<SessionToken> {

        public SessionToken(Session session)
            : base(new SessionId(session.SessionId)) {
            Debug.WriteLine("SessionToken.Constructor");

            var reachableApps = session.Apps
                .Select(a => new AppToken(a))
                .OrderBy(at => at.Id);

            var unreachableApps = session.UnreachableProcessIds
                .Select(id => AppToken.Unreachable(id))
                .OrderBy(at => at.Id);

            Apps = reachableApps
                .Concat(unreachableApps)
                .ToImmutableArray();

            int? topMostId = session.TopMost?.AsProcess()?.Id;
            if (topMostId.HasValue)
                ActiveApp = Apps.SingleOrDefault(a => a.Id.ProcessId == topMostId);

            int? primaryId = session.Primary?.AsProcess()?.Id;
            if (primaryId.HasValue)
                Primary = Apps.SingleOrDefault(a => a.Id.ProcessId == primaryId);

            //   Debug.WriteLine("SessionToken.Constructor (end)");
            //  Debug.WriteLine("");
        }

        public IEnumerable<AppToken> Apps { get; }

        public AppToken ActiveApp { get; }

        public AppToken Primary { get; }

        #region Equality

        public bool Equals(SessionToken other) => base.Equals(other)
            && Apps.SequenceEqual(other.Apps)
            && Equals(ActiveApp, other.ActiveApp)
            && Equals(Primary, other.Primary);

        public override bool Equals(object obj) => Equals(obj as SessionToken);

        public bool Matches(SessionToken other) => base.Equals(other);

        #endregion
    }
}
