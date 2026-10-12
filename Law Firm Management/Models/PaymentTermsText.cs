namespace Law_Firm_Management.Models
{
    /// <summary>
    /// Display and comparison rules for payment terms. Terms are stored as a number of days.
    /// Blank (NULL) and 0 both mean Due on Receipt, so they are treated as the same value.
    /// </summary>
    public static class PaymentTermsText
    {
        /// <summary>Text shown on screens and in the audit history.</summary>
        public static string Format(short? terms)
        {
            if (terms == null || terms.Value == 0)
                return "Due on Receipt";

            return terms.Value == 1 ? "1 day" : $"{terms.Value} days";
        }

        /// <summary>True when both values describe the same terms (blank and 0 are equal).</summary>
        public static bool AreSame(short? a, short? b)
        {
            return (a ?? 0) == (b ?? 0);
        }
    }
}
