namespace SmartOCR.Utilities
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using SmartOCR.Word;

    /// <summary>
    /// Represents grid-like structure, that contains pairs of line mapping and <see cref="WordTable"/> collection, broken into evenly distributed segments.
    /// </summary>
    public class GridStructure
    {
        /// <summary>
        /// Matrix of pairs (line mapping and tables).
        /// </summary>
        private Tuple<LineMapping, List<WordTable>>[][] grid;
        private LineMapping source = new LineMapping();
        private List<WordTable> tables = new List<WordTable>();
        private int singleSegmentRowSize;
        private decimal singleSegmentColumnSize;

        /// <summary>
        /// Initializes a new instance of the <see cref="GridStructure"/> class.
        /// </summary>
        /// <param name="lineMapping">A source line mapping.</param>
        /// <param name="tableCollection">A collection of <see cref="WordTable"/>.</param>
        public GridStructure(LineMapping lineMapping, List<WordTable> tableCollection)
        {
            this.InitializeFields(lineMapping, tableCollection);
            this.InitializeGrid();
            this.PopulateGrid();
            this.AddTablesToGrid();
        }

        /// <summary>
        /// Gets or sets grid size for every instance of <see cref="GridStructure"/>.
        /// </summary>
        public static int StaticSize { get; set; } = 3;

        /// <summary>
        /// Gets grid size.
        /// </summary>
        public int Size { get; private set; }

        /// <summary>
        /// Gets line mapping in specified grid gridSegment.
        /// </summary>
        /// <param name="gridRow">Row index.</param>
        /// <param name="gridColumn">Column index.</param>
        /// <returns>A pair of <see cref="LineMapping"/> instance and <see cref="WordTable"/> collection.</returns>
        public Tuple<LineMapping, List<WordTable>> this[int gridRow, int gridColumn] =>
            this.AreIndexersValid(gridRow, gridColumn)
                ? this.grid[gridRow][gridColumn]
                : null;

        /// <summary>
        /// Gets line mapping in specified grid gridSegment.
        /// </summary>
        /// <param name="gridCoordinates">A tuple, containing grid coordinates.</param>
        /// <returns>A pair of <see cref="LineMapping"/> instance and <see cref="WordTable"/> collection.</returns>
        public Tuple<LineMapping, List<WordTable>> this[Tuple<int, int> gridCoordinates] => this[gridCoordinates.Item1, gridCoordinates.Item2];

        /// <inheritdoc/>
        public override string ToString()
        {
            try
            {
                StringBuilder builder = new StringBuilder();
                for (int row = 0; row < this.Size; row++)
                {
                    ParagraphContainer firstContainerInSegment = this[row, 0].Item1.Values.Where(item => item.Count > 0).First().First();
                    builder.AppendFormat("({0}, {1})", firstContainerInSegment.HorizontalLocation, firstContainerInSegment.VerticalLocation);
                    for (int column = 1; column < this.Size; column++)
                    {
                        firstContainerInSegment = this[row, column].Item1.Values.Where(item => item.Count > 0).First().First();
                        builder.Append("-----").AppendFormat("({0}, {1})", firstContainerInSegment.HorizontalLocation, firstContainerInSegment.VerticalLocation);
                    }

                    builder.AppendLine();
                }

                return builder.ToString();
            }
            catch (InvalidOperationException)
            {
                return base.ToString();
            }
        }

        private static bool DoesGridSegmentContainsTableAnchor(LineMapping gridSegment, ParagraphContainer anchor)
        {
            return gridSegment.Any(
                lineContent => lineContent.Value
                    .Any(paragraphContainer => paragraphContainer.HorizontalLocation == anchor.HorizontalLocation &&
                                               paragraphContainer.VerticalLocation == anchor.VerticalLocation));
        }

        private bool AreIndexersValid(int gridRow, int gridColumn)
        {
            return gridRow >= 0 && gridRow < this.Size &&
                   gridColumn >= 0 && gridRow < this.Size;
        }

        private void InitializeFields(LineMapping lineMapping, List<WordTable> tableCollection)
        {
            this.source = lineMapping;
            this.Size = StaticSize;
            this.tables = tableCollection;
            this.InitializeSegmentDimensions();
        }

        private void InitializeSegmentDimensions()
        {
            this.singleSegmentRowSize = (this.source.Count / this.Size) + 1;
            this.singleSegmentColumnSize = this.GetColumnSize() / this.Size;
        }

        private void InitializeGrid()
        {
            this.grid = new Tuple<LineMapping, List<WordTable>>[this.Size][];
            for (int gridRow = 0; gridRow < this.Size; gridRow++)
            {
                this.InitializeGridRow(gridRow);
            }
        }

        private void InitializeGridRow(int gridRow)
        {
            this.grid[gridRow] = new Tuple<LineMapping, List<WordTable>>[this.Size];
            this.AddEmptyMappingsToGridRow(gridRow);
        }

        private void AddEmptyMappingsToGridRow(int gridRow)
        {
            for (int gridColumn = 0; gridColumn < this.Size; gridColumn++)
            {
                this.grid[gridRow][gridColumn] = new Tuple<LineMapping, List<WordTable>>(new LineMapping(), new List<WordTable>());
            }
        }

        private void PopulateGrid()
        {
            for (int gridRow = 0; gridRow < this.Size; gridRow++)
            {
                this.PopulateGridRow(gridRow);
            }
        }

        private void AddTablesToGrid()
        {
            foreach (var table in this.tables)
            {
                this.AddSingleTableToGrid(table);
            }
        }

        private void AddSingleTableToGrid(WordTable table)
        {
            for (int i = 0; i < this.Size; i++)
            {
                for (int j = 0; j < this.Size; j++)
                {
                    var segment = this.grid[i][j];
                    if (!DoesGridSegmentContainsTableAnchor(segment.Item1, table.Anchor))
                    {
                        continue;
                    }

                    segment.Item2.Add(table);
                    return;
                }
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
            var splitParagraphCollection = this.SplitParagraphCollection(this.source[lineCount]);
            this.PopulateColumnsInGridRow(gridRow, lineCount, splitParagraphCollection);
        }

        private void PopulateColumnsInGridRow(int gridRow, int lineCount, List<ParagraphContainer>[] splitParagraphCollection)
        {
            for (int gridColumn = 0; gridColumn < this.Size; gridColumn++)
            {
                this.AddParagraphsAndSort(
                    this.grid[gridRow][gridColumn].Item1,
                    lineCount,
                    splitParagraphCollection[gridColumn]);
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
            return this.DistributeSourceParagraphs(paragraphs, this.InitializeSplitParagraphs());
        }

        private List<ParagraphContainer>[] DistributeSourceParagraphs(List<ParagraphContainer> sourceParagraphs, List<ParagraphContainer>[] splitParagraphs)
        {
            foreach (var item in sourceParagraphs)
            {
                splitParagraphs[this.GetGridColumnForParagraph(item)].Add(item);
            }

            return splitParagraphs;
        }

        private List<ParagraphContainer>[] InitializeSplitParagraphs()
        {
            return this.InitializeSplitParagraphsWithEmptyLists(new List<ParagraphContainer>[this.Size]);
        }

        private List<ParagraphContainer>[] InitializeSplitParagraphsWithEmptyLists(List<ParagraphContainer>[] splitParagraphs)
        {
            for (int i = 0; i < this.Size; i++)
            {
                splitParagraphs[i] = new List<ParagraphContainer>();
            }

            return splitParagraphs;
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
