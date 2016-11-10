using System;

namespace ExcelBrowser.Model {

    public abstract class Token<TId> : IEquatable<Token<TId>> {

        protected Token(TId id) {
            Requires.NotNull(id, nameof(id));
            Id = id;
        }

        public TId Id { get; }

        #region Equality

        public bool Equals(Token<TId> other) => Equals(Id, other.Id);
        
        public override bool Equals(object obj) => Equals(obj as Token<TId>);

        public override int GetHashCode() => Id.GetHashCode();

        #endregion

        public override string ToString() => Id.ToString();
    }
}
