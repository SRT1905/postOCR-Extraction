namespace SmartOCR
{
    using System.Collections.Generic;

    /// <summary>
    /// Represents mapping of multiple Word paragraphs to their line location on document page.
    /// </summary>
    public class LineMapping : SortedDictionary<int, List<ParagraphContainer>>
    {
        /// <summary>
        /// Changes specific key in dictionary.
        /// If dictionary does not contain original key, then new item is created with target key and empty <see cref="List{ParagraphContainer}"/>.
        /// </summary>
        /// <param name="fromKey">Source key.</param>
        /// <param name="toKey">Target key.</param>
        public void RenameKey(int fromKey, int toKey)
        {
            if (!this.ContainsKey(fromKey))
            {
                this.AddNewList(toKey);
                return;
            }

            this.AddExistingParagraphs(toKey, this[fromKey]);
            this.Remove(fromKey);
        }

        private void AddExistingParagraphs(int toKey, List<ParagraphContainer> paragraphs)
        {
            if (!this.ContainsKey(toKey))
            {
                this.Add(toKey, new List<ParagraphContainer>());
            }

            this[toKey] = paragraphs;
        }

        private void AddNewList(int toKey)
        {
            if (!this.ContainsKey(toKey))
            {
                this.Add(toKey, new List<ParagraphContainer>());
            }
            else
            {
                this[toKey] = new List<ParagraphContainer>();
            }
        }
    }
}
