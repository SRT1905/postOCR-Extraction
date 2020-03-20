namespace SmartOCR.Utilities
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Represents collection of <see cref="GridStructure"/> instances, accessible by page number.
    /// </summary>
    [Serializable]
    public class GridCollection : Dictionary<int, GridStructure>
    {
    }
}
