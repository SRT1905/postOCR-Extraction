namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Represents grid-like structure, that contains line mapping, broken into evenly distrubuted segments.
    /// </summary>
    public class GridStructure
    {
        private LineMapping[][] grid;
        private LineMapping source;
        private int singleSegmentRowSize;
        private decimal singleSegmentColumnSize;

        /// <summary>
        /// Initializes a new instance of the <see cref="GridStructure"/> class.
        /// </summary>
        /// <param name="lineMapping">A source line mapping.</param>
        /// <param name="size">Max number of segments (both rows and columns).</param>
        public GridStructure(LineMapping lineMapping, int size)
        {
            this.InitializeFields(lineMapping, size);
            this.InitializeGrid();
            this.PopulateGrid();
        }

        /// <summary>
        /// Gets grid size.
        /// </summary>
        public int Size { get; private set; }

        /// <summary>
        /// Gets line mapping in specified grid segment.
        /// </summary>
        /// <param name="gridRow">Row index.</param>
        /// <param name="gridColumn">Column index.</param>
        /// <returns>An <see cref="LineMapping"/> instance.</returns>
        public LineMapping this[int gridRow, int gridColumn]
        {
            get
            {
                return this.AreIndexersValid(gridRow, gridColumn)
                    ? this.grid[gridRow][gridColumn]
                    : null;
            }
        }

        private bool AreIndexersValid(int gridRow, int gridColumn)
        {
            return gridRow >= 0 && gridRow < this.Size &&
                   gridColumn >= 0 && gridRow < this.Size;
        }

        private void InitializeFields(LineMapping lineMapping, int size)
        {
            this.source = lineMapping;
            this.Size = size;
            this.InitializeSegmentDimensions();
        }

        private void InitializeSegmentDimensions()
        {
            this.singleSegmentRowSize = (this.source.Count / this.Size) + 1;
            this.singleSegmentColumnSize = this.GetColumnSize() / this.Size;
        }

        private void InitializeGrid()
        {
            this.grid = new LineMapping[this.Size][];
            for (int gridRow = 0; gridRow < this.Size; gridRow++)
            {
                this.InitializeGridRow(gridRow);
            }
        }

        private void InitializeGridRow(int gridRow)
        {
            this.grid[gridRow] = new LineMapping[this.Size];
            this.AddEmptyMappingsToGridRow(gridRow);
        }

        private void AddEmptyMappingsToGridRow(int gridRow)
        {
            for (int gridColumn = 0; gridColumn < this.Size; gridColumn++)
            {
                this.grid[gridRow][gridColumn] = new LineMapping();
            }
        }

        private void PopulateGrid()
        {
            for (int gridRow = 0; gridRow < this.Size; gridRow++)
            {
                this.PopulateGridRow(gridRow);
            }
        }

        private void PopulateGridRow(int gridRow)
        {
            for (int lineCount = gridRow * this.singleSegmentRowSize; lineCount < (gridRow + 1) * this.singleSegmentRowSize; lineCount++)
            {
                this.TryAddLineContentToGrid(gridRow, lineCount);
            }
        }

        private void TryAddLineContentToGrid(int gridRow, int lineCount)
        {
            if (this.source.ContainsKey(lineCount))
            {
                this.AddSingleLineDataToGrid(gridRow, lineCount);
            }
        }

        private void AddSingleLineDataToGrid(int gridRow, int lineCount)
        {
            var splittedParagraphCollection = this.SplitParagraphCollection(this.source[lineCount]);
            this.PopulateColumnsInGridRow(gridRow, lineCount, splittedParagraphCollection);
        }

        private void PopulateColumnsInGridRow(int gridRow, int lineCount, List<ParagraphContainer>[] splittedParagraphCollection)
        {
            for (int gridColumn = 0; gridColumn < this.Size; gridColumn++)
            {
                this.AddParagraphsAndSort(
                    this.grid[gridRow][gridColumn],
                    lineCount,
                    splittedParagraphCollection[gridColumn]);
            }
        }

        private void AddParagraphsAndSort(LineMapping gridSegment, int lineNumber, List<ParagraphContainer> paragraphs)
        {
            gridSegment.Add(lineNumber, paragraphs);
            gridSegment[lineNumber].Sort();
        }

        private decimal GetColumnSize()
        {
            return this.source
                       .Select(keyValuePair => keyValuePair.Value.Select(item => item.HorizontalLocation).Max())
                       .Max();
        }

        private List<ParagraphContainer>[] SplitParagraphCollection(List<ParagraphContainer> paragraphs)
        {
            return this.DistributeSourceParagraphs(paragraphs, this.InitializeSplittedParagraphs());
        }

        private List<ParagraphContainer>[] DistributeSourceParagraphs(List<ParagraphContainer> sourceParagraphs, List<ParagraphContainer>[] splittedParagraphs)
        {
            foreach (ParagraphContainer item in sourceParagraphs)
            {
                splittedParagraphs[this.GetGridColumnForParagraph(item)].Add(item);
            }

            return splittedParagraphs;
        }

        private List<ParagraphContainer>[] InitializeSplittedParagraphs()
        {
            return this.InitializeSplittedParagraphsWithEmptyLists(new List<ParagraphContainer>[this.Size]);
        }

        private List<ParagraphContainer>[] InitializeSplittedParagraphsWithEmptyLists(List<ParagraphContainer>[] splittedParagraphs)
        {
            for (int i = 0; i < this.Size; i++)
            {
                splittedParagraphs[i] = new List<ParagraphContainer>();
            }

            return splittedParagraphs;
        }

        private int GetGridColumnForParagraph(ParagraphContainer item)
        {
            int index = Convert.ToInt32(Math.Floor(item.HorizontalLocation / this.singleSegmentColumnSize));
            return index >= this.Size
                ? this.Size - 1
                : index;
        }
    }
}
