using OfficeDesk.Api.Services;
using Xunit;

namespace OfficeDesk.Api.Tests;

public class MailSummarizerTests
{
    [Fact]
    public void PicksTheRequestAndDropsTheGreetingAndSignature()
    {
        var body = """
            السلام عليكم،
            نرجو اعتماد عرض السعر قبل يوم الخميس.
            المرفق فيه التفاصيل.
            مع التحية
            سارة
            """;

        var summary = MailSummarizer.Summarize("عرض السعر", body);

        Assert.Contains("اعتماد عرض السعر", summary);
        Assert.DoesNotContain("السلام عليكم", summary);
        Assert.DoesNotContain("مع التحية", summary);
    }

    [Fact]
    public void IgnoresTheQuotedReply()
    {
        var body = """
            Please confirm the invoice today.
            Thanks
            -----Original Message-----
            From: Omar
            Can you send the old file?
            """;

        var summary = MailSummarizer.Summarize("Invoice", body);

        Assert.Contains("confirm the invoice", summary);
        Assert.DoesNotContain("old file", summary);
    }

    [Fact]
    public void SkipsLogoLinksAndKeepsTheRequest()
    {
        var body = "<https://cdn.example/logo.png> <https://twitter.com/hr> نرجو اعتماد الإجازة اليوم.";

        var summary = MailSummarizer.Summarize("إجازة", body);

        Assert.Contains("اعتماد الإجازة", summary);
        Assert.DoesNotContain("http", summary);
    }

    [Fact]
    public void ShortensAnEnglishNoticeTheSameWay()
    {
        var body = """
            Dear colleagues,
            Regarding the shift schedule, please note that the 7:00 a.m. to 4:00 p.m. shift lasts for one month only. After that month, the employee moves to the core shift from 8:00 a.m. to 5:00 p.m. for two months.
            Employees who were on the 7:00 a.m. shift in September will move next week, for October and November.
            Thank you for your cooperation.
            Best regards,
            """;

        var summary = MailSummarizer.Summarize("Shift schedule", body);

        Assert.StartsWith("Please note", summary);
        Assert.Contains("7:00 a.m. to 4:00 p.m.", summary);
        Assert.Contains("one month only", summary);
        Assert.Contains("After that month, the employee moves", summary);
        Assert.Contains("two months", summary);
        Assert.Contains("September", summary);
        Assert.Contains("October", summary);
        Assert.DoesNotContain("Dear", summary);
        Assert.DoesNotContain("regards", summary);
    }

    [Fact]
    public void DropsTheEnglishGreetingAroundACode()
    {
        var body = """
            Your client number: K0000000001
            Dear Sara Nasser
            Please enter the following verification code: 100200
            If you did not request this code, you can ignore this email.
            Best regards
            """;

        var summary = MailSummarizer.Summarize("Verification code", body);

        Assert.Contains("verification code: 100200", summary);
        Assert.DoesNotContain("Dear", summary);
        Assert.DoesNotContain("regards", summary);
    }

    [Fact]
    public void KeepsTheShiftDecision()
    {
        var body = """
            عناية/ منسوبي شركة إجادة المهنية المحترمين
            السلام عليكم ورحمة الله وبركاته،،
            بخصوص تنظيم الشفتات، نود التوضيح بأن شفت الساعة 7:00 صباحًا إلى الساعة 4:00 مساءً يكون لمدة شهر واحد فقط، وبعد انتهاء الشهر ينتقل الموظف إلى الشفت الأساسي من الساعة 8:00 صباحًا إلى الساعة 5:00 مساءً لمدة شهرين.
            وعليه، بالنسبة للموظفين الذين كانوا على شفت الساعة 7:00 صباحًا خلال شهر سبتمبر، سيكون دوامهم اعتبارًا من الأسبوع القادم على الشفت الأساسي من الساعة 8:00 صباحًا إلى الساعة 5:00 مساءً، وذلك لمدة شهرين (أكتوبر ونوفمبر).
            وبعد انتهاء فترة الشهرين، وفي حال رغبة الموظف بالعودة إلى شفت الساعة 7:00 صباحًا، يمكنه العودة إليه وفقًا للتنظيم المعتمد.
            ويأتي هذا التنظيم بهدف تدوير الشفتات بين الموظفين بشكل منتظم.
            شاكرين لكم تعاونكم والتزامكم.
            وتفضلوا بقبول وافر التحية والتقدير،،
            """;

        var summary = MailSummarizer.Summarize("توضيح بشأن نظام الشفتات", body);

        Assert.Contains("7:00", summary);
        Assert.Contains("شهر واحد", summary);
        Assert.Contains("سبتمبر", summary);
        Assert.Contains("الأسبوع القادم", summary);
        Assert.Contains("أكتوبر", summary);
        Assert.Contains("العودة", summary);
        Assert.DoesNotContain("السلام عليكم", summary);
        Assert.DoesNotContain("المحترمين", summary);
    }

    [Fact]
    public void SkipsALeadingDisclaimerAndNewsletterNoise()
    {
        var body = """
            This email and any attachment are confidential.
            View in browser
            Please approve the contract by Thursday.
            Unsubscribe
            """;

        var summary = MailSummarizer.Summarize("Contract", body);

        Assert.Contains("approve the contract", summary);
        Assert.DoesNotContain("confidential", summary);
        Assert.DoesNotContain("Unsubscribe", summary);
    }

    [Fact]
    public void KeepsTheMessageWhenItStartsWithMin()
    {
        var body = """
            من الموظفين من كان على الشفت الصباحي يكمل حتى يوم الخميس.
            نرجو اعتماد الجدول اليوم.
            """;

        var summary = MailSummarizer.Summarize("الجدول", body);

        Assert.Contains("اعتماد الجدول", summary);
        Assert.Contains("يوم الخميس", summary);
    }

    [Fact]
    public void UsesTheSubjectWhenTheBodyIsEmpty()
    {
        Assert.Equal("موعد الاجتماع", MailSummarizer.Summarize("موعد الاجتماع", "   "));
    }
}
