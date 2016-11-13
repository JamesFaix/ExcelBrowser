using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Runtime.Serialization;

#pragma warning disable CS0659 //Does not need to override GetHashCode because base class implementation is sufficient.

namespace ExcelBrowser.Model {

    [DataContract]
    public class SessionToken : Token<SessionId>, IEquatable<SessionToken> {

        public SessionToken(SessionId id, IEnumerable<AppToken> apps, AppId primaryAppId) 
            : base(id) {
            Requires.NotNull(apps, nameof(apps));
            Requires.Rule(primaryAppId == null || apps.Select(a => a.Id).Contains(primaryAppId),
                "PrimaryApp ID must be null or belong to one of the given apps.");

            Apps = apps.ToImmutableArray();
            PrimaryAppId = primaryAppId;
        }
        
        [DataMember(Order = 2)]
        public IEnumerable<AppToken> Apps { get; }

        [DataMember(Order = 3)]
        public AppId PrimaryAppId { get; }

        #region Equality

        public bool Equals(SessionToken other) => base.Equals(other)
            && Apps.SequenceEqual(other.Apps)
            && Equals(PrimaryAppId, other.PrimaryAppId);

        public override bool Equals(object obj) => Equals(obj as SessionToken);

        public bool Matches(SessionToken other) => base.Equals(other);

        #endregion

        public override string ToString() => Serializer.Serialize(this);
    }
}
