using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Runtime.Serialization;

#pragma warning disable CS0659 //Does not need to override GetHashCode because base class implementation is sufficient.

namespace ExcelBrowser.Model {

    [DataContract]
    public class SessionToken : Token<SessionId>, IEquatable<SessionToken> {

        internal SessionToken(SessionId id, IEnumerable<AppToken> apps, 
            AppId activeAppId, AppId primaryAppId) : base(id) {
            Requires.NotNull(apps, nameof(apps));

            Apps = apps.ToImmutableArray();

            if (activeAppId != null) {
                try {
                    ActiveApp = Apps.Single(a => Equals(a.Id, activeAppId));
                }
                catch (InvalidOperationException x)
                when (x.Message.StartsWith("Sequence contains no elements")) {
                    throw new InvalidOperationException("ActiveApp ID not found in apps collection.", x);
                }
            }

            if (primaryAppId != null) {
                try {
                    PrimaryApp = Apps.Single(a => Equals(a.Id, primaryAppId));
                }
                catch (InvalidOperationException x)
                when (x.Message.StartsWith("Sequence contains no elements")) {
                    throw new InvalidOperationException("PrimaryApp ID not found in apps collection.", x);
                }
            }
        }
        
        [DataMember(Order = 2)]
        public IEnumerable<AppToken> Apps { get; }

        [DataMember(Order = 3)]
        public AppToken ActiveApp { get; }

        [DataMember(Order = 4)]
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
