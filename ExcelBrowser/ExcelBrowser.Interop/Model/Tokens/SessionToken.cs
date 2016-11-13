using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Serialization;
using ExcelBrowser.Interop;

#pragma warning disable CS0659 //Does not need to override GetHashCode because base class implementation is sufficient.

namespace ExcelBrowser.Model {

    [DataContract]
    public class SessionToken : Token<SessionId>, IEquatable<SessionToken> {

        public SessionToken(Session session)
            : base(new SessionId(session.SessionId)) {
          //  Debug.WriteLine("SessionToken.Constructor");

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
                PrimaryApp = Apps.SingleOrDefault(a => a.Id.ProcessId == primaryId);

            //   Debug.WriteLine("SessionToken.Constructor (end)");
            //  Debug.WriteLine("");
        }

        [DataMember(Order = 1)]
        public IEnumerable<AppToken> Apps { get; }

        [DataMember(Order = 2)]
        public AppToken ActiveApp { get; }

        [DataMember(Order = 3)]
        public AppToken PrimaryApp { get; }

        #region Equality

        public bool Equals(SessionToken other) => base.Equals(other)
            && Apps.SequenceEqual(other.Apps)
            && Equals(ActiveApp, other.ActiveApp)
            && Equals(PrimaryApp, other.PrimaryApp);

        public override bool Equals(object obj) => Equals(obj as SessionToken);

        public bool Matches(SessionToken other) => base.Equals(other);

        #endregion

        public override string ToString() => Serializer.Serialize(this);
    }
}
