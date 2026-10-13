using System.Text.RegularExpressions;

namespace OfficeDesk.Api.Services;

public static partial class MailSummarizer
{
    public static string Summarize(string? subject, string? body)
    {
        var clean = Clean(body);
        if (clean.Length < 12)
            return Clip(subject);

        var sentences = Sentences(clean);
        var scored = sentences
            .Select((sentence, index) => (sentence, index, score: Score(sentence)))
            .ToList();
        var picked = TakeFacts(scored.Where(item => item.score > 0).OrderBy(item => item.index));
        if (picked.Count == 0)
            picked = TakeFacts(scored.Where(item => item.score >= 0).OrderBy(item => item.index).Take(2));

        var summary = string.Join(" ", picked);
        if (string.IsNullOrWhiteSpace(summary))
            summary = Clip(subject);
        return Capitalize(summary);
    }

    static string Clean(string? body)
    {
        if (string.IsNullOrWhiteSpace(body))
            return "";

        var source = Link().Replace(body, " ");
        source = Tag().Replace(source, " ");
        var kept = new List<string>();
        foreach (var raw in source.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
        {
            var line = raw.Trim();
            if (line.Length == 0)
                continue;
            if (line.StartsWith('>'))
                continue;
            if (IsNoise(line) || IsDisclaimer(line))
                continue;
            if (IsQuote(line) || (kept.Count > 0 && IsSignature(line)))
                break;
            line = DearSalutation().Replace(line, " ").Trim();
            line = Regarding().Replace(line, "").Trim();
            line = Bikhusoos().Replace(line, "").Trim();
            if (line.Length == 0 || IsGreeting(line))
                continue;
            kept.Add(line);
        }

        var text = Spaces().Replace(string.Join(" ", kept), " ").Trim();
        return TrailingHello().Replace(text, "").Trim();
    }

    static List<string> Sentences(string text)
    {
        var parts = SentenceSplit().Split(Protect(text))
            .Select(part => part.Replace('\u0001', '.'))
            .Select(part => part.Trim().TrimEnd('،', ',', ';', '؛', ' '))
            .Where(part => part.Length >= 8 && part.Count(char.IsLetter) >= 8)
            .ToList();
        return parts.Count == 0 ? [text] : parts;
    }

    static int Score(string sentence)
    {
        var text = sentence.ToLowerInvariant();
        var score = 0;
        if (sentence.Contains('?') || sentence.Contains('؟'))
            score += 4;
        if (Asks().IsMatch(text))
            score += 5;
        if (Timing().IsMatch(text))
            score += 3;
        if (Digit().IsMatch(sentence))
            score += 1;
        if (Noise().IsMatch(text))
            score -= 4;
        if (sentence.Length < 40 && Greeting().IsMatch(text))
            score -= 4;
        return score;
    }

    static List<string> TakeFacts(IEnumerable<(string sentence, int index, int score)> items)
    {
        var chosen = new List<string>();
        foreach (var item in items)
        {
            if (chosen.Count == 4)
                break;
            if (chosen.Any(existing => existing.Contains(item.sentence, StringComparison.Ordinal) || item.sentence.Contains(existing, StringComparison.Ordinal)))
                continue;
            var next = chosen.Count == 0 ? item.sentence : string.Join(" ", chosen) + " " + item.sentence;
            if (next.Length > 700 && chosen.Count > 0)
                break;
            chosen.Add(item.sentence);
        }
        return chosen;
    }

    static bool IsNoise(string line)
    {
        var text = line.ToLowerInvariant();
        return line.Length < 80 && Noise().IsMatch(text);
    }

    static bool IsQuote(string line) =>
        line.StartsWith("-----", StringComparison.Ordinal)
        || line.StartsWith("_____", StringComparison.Ordinal)
        || line.StartsWith("ــــ", StringComparison.Ordinal)
        || line.Contains("Original Message", StringComparison.OrdinalIgnoreCase)
        || line.Contains("Forwarded message", StringComparison.OrdinalIgnoreCase)
        || (line.StartsWith("On ", StringComparison.OrdinalIgnoreCase) && line.Contains("wrote:", StringComparison.OrdinalIgnoreCase))
        || (line.StartsWith("From:", StringComparison.OrdinalIgnoreCase) && line.Length < 100)
        || (line.StartsWith("من:", StringComparison.Ordinal) && line.Length < 50)
        || (line.Contains("كتب:", StringComparison.Ordinal) && line.Length < 80);

    static string Protect(string text) =>
        Abbreviation().Replace(text, match => match.Value.Replace('.', '\u0001'));

    static bool IsGreeting(string line)
    {
        var text = Spaces().Replace(line, " ").Trim().TrimEnd('.', '،', ',', '!', ':', ' ').ToLowerInvariant();
        if (text.Length is 0 or > 80)
            return false;
        return text is "hi" or "hello" or "hey" or "good morning" or "good afternoon" or "good evening"
            or "dear all" or "dear team" or "dear colleagues" or "dear sir" or "dear madam"
            or "السلام عليكم" or "السلام عليكم ورحمة الله وبركاته" or "صباح الخير" or "مساء الخير"
            || text.Contains("المحترمين", StringComparison.Ordinal)
            || text.StartsWith("dear ", StringComparison.Ordinal)
            || text.StartsWith("hello ", StringComparison.Ordinal)
            || text.StartsWith("hi ", StringComparison.Ordinal)
            || text.StartsWith("السلام عليكم", StringComparison.Ordinal)
            || text.StartsWith("مرحبا", StringComparison.Ordinal);
    }

    static bool IsSignature(string line)
    {
        if (line.Length > 60)
            return false;
        var text = line.Trim().TrimEnd('.', '،', ',', '!', ' ').ToLowerInvariant();
        return text is "best regards" or "kind regards" or "warm regards" or "regards" or "sincerely"
            or "yours sincerely" or "thanks" or "thank you" or "thanks and regards" or "many thanks"
            or "مع التحية" or "مع خالص التحية" or "وتقبلوا فائق الاحترام" or "تحياتي" or "خالص التحية"
            || text.StartsWith("sent from", StringComparison.Ordinal)
            || text.StartsWith("best regards", StringComparison.Ordinal)
            || text.StartsWith("kind regards", StringComparison.Ordinal);
    }

    static bool IsDisclaimer(string line)
    {
        var text = line.ToLowerInvariant();
        return text.Contains("this email and any")
            || text.Contains("confidential")
            || (text.Contains("هذه الرسالة") && text.Contains("سر"));
    }

    static string Capitalize(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;
        return char.ToUpper(text[0]) + text[1..];
    }

    static string Clip(string? value, int max = 180)
    {
        var text = Spaces().Replace(value ?? "", " ").Trim();
        if (text.Length <= max)
            return text;
        var cut = text.LastIndexOf(' ', Math.Min(max, text.Length - 1));
        if (cut < 40)
            cut = max;
        return text[..cut].TrimEnd('،', ',', '.', '؛', ' ') + "…";
    }

    [GeneratedRegex(@"<https?://[^>\s]+>|https?://\S+|www\.\S+", RegexOptions.IgnoreCase)]
    private static partial Regex Link();

    [GeneratedRegex(@"<[^>]+>")]
    private static partial Regex Tag();

    [GeneratedRegex(@"(?i)\b(?:a\.m|p\.m|mr|mrs|ms|dr|e\.g|i\.e)\.")]
    private static partial Regex Abbreviation();

    [GeneratedRegex(@"(?i)\bdear\s+[a-z]+(?:\s+[a-z]+){0,2}\b[,:]?\s*")]
    private static partial Regex DearSalutation();

    [GeneratedRegex(@"(?i)\s+\b(?:hi|hello|hey)\b\s*$")]
    private static partial Regex TrailingHello();

    [GeneratedRegex(@"\s+")]
    private static partial Regex Spaces();

    [GeneratedRegex(@"(?<=[\.!\?؟])\s+")]
    private static partial Regex SentenceSplit();

    [GeneratedRegex(@"(?i)^regarding\b[^,]{0,80},\s*")]
    private static partial Regex Regarding();

    [GeneratedRegex(@"^بخصوص\b[^،]{0,60}،\s*")]
    private static partial Regex Bikhusoos();

    [GeneratedRegex(@"يرجى|نرجو|أرجو|ارجو|فضلا|لو سمحت|please|kindly|could you|can you|نحتاج|مطلوب|action required|اعتماد|موافقة|approve|confirm|تأكيد")]
    private static partial Regex Asks();

    [GeneratedRegex(@"اليوم|غدا|غدًا|قبل |موعد|deadline|\bby\b|asap|عاجل|فوري|أسبوع|الاسبوع|الأسبوع|شهر|سبتمبر|أكتوبر|اكتوبر|نوفمبر|ديسمبر|يناير|الأحد|الاثنين|الثلاثاء|الأربعاء|الخميس|الجمعة|السبت|\bweek\b|\bmonth\b|september|october|november|december|january|monday|tuesday|wednesday|thursday|friday|saturday|sunday|tomorrow")]
    private static partial Regex Timing();

    [GeneratedRegex(@"unsubscribe|view in browser|click here|follow us|all rights reserved|إلغاء الاشتراك|اضغط هنا|عرض خاص")]
    private static partial Regex Noise();

    [GeneratedRegex(@"\d")]
    private static partial Regex Digit();

    [GeneratedRegex(@"السلام عليكم|مرحبا|مرحبًا|^dear |^hello|^hi |good morning|صباح الخير")]
    private static partial Regex Greeting();
}
