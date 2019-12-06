using System;

namespace SmartOCR
{
    internal class SimilarityDescription
    {
        public double Ratio;
        public string Value;
        private const double similarity_ratio_threshold = 0.66;
        private const byte string_length_offset = 1;

        public SimilarityDescription(string found_string, string check_string)
        {
            if (found_string.Length - string_length_offset <= check_string.Length &&
                found_string.Length + string_length_offset >= check_string.Length)
            {
                Ratio = GetStringSimilarity(found_string, check_string);
                Value = found_string;
            }
        }

        public bool CheckStringSimilarity()
        {
            return Ratio >= similarity_ratio_threshold;
        }

        private double GetStringSimilarity(string left_string, string right_string)
        {
            string short_string;
            string long_string;
            if (left_string.Length < right_string.Length)
            {
                short_string = left_string;
                long_string = right_string;
            }
            else
            {
                short_string = right_string;
                long_string = left_string;
            }
            return long_string.ToLower() == short_string.ToLower()
                ? 1
                : (long_string.Length - ComputeLevensteinDistance(long_string, short_string)) / long_string.Length;
        }

        private double ComputeLevensteinDistance(string long_string, string short_string)
        {
            int long_length = long_string.Length;
            int short_length = short_string.Length;

            int[,] costs = new int[long_length + 1, short_length + 1];

            if (long_length == 0)
            {
                return short_length;
            }

            if (short_length == 0)
            {
                return long_length;
            }

            for (int i = 0; i <= long_length; costs[i, 0] = i++) ;
            for (int j = 0; j <= short_length; costs[0, j] = j++) ;
            for (int i = 1; i <= long_length; i++)
            {
                for (int j = 1; j <= short_length; j++)
                {
                    int cost = (short_string[j - 1] == long_string[i - 1]) ? 0 : 1;
                    costs[i, j] = Math.Min(
                        Math.Min(
                            costs[i - 1, j] + 1,
                            costs[i, j - 1] + 1),
                        costs[i - 1, j - 1] + cost);
                }
            }
            return costs[long_length, short_length];
        }
    }
}