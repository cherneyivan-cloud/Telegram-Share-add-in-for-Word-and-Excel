namespace TelegramShareAddin
{
    /// <summary>Идентификаторы элементов ленты — единственный источник правды.</summary>
    internal static class RibbonIds
    {
        public const string Tab = "TelegramShareTab";
        public const string SendVia = "TelegramShareSendViaButton";
        public const string SendMessage = "TelegramShareButton";
        public const string Attachments = "TelegramShareAttachmentsButton";
        public const string Print = "TelegramSharePrintButton";
        public const string Max = "TelegramShareMaxButton";
        public const string PinQat = "TelegramSharePinQatButton";
        public const string Update = "TelegramShareUpdateButton";
    }

    /// <summary>Подписи кнопок (по ним же ищет кнопки закрепление в панели быстрого доступа).</summary>
    internal static class RibbonLabels
    {
        public const string SendMessage = "Отправить в Telegram";
        public const string Max = "Отправить в MAX";
    }

    /// <summary>Разметка ленты Office, общая для Word, Excel и Outlook.</summary>
    internal static class RibbonDefinition
    {
        // Кнопка «Вложения» нужна только Outlook — в Word/Excel ленте её нет.
        private const string AttachmentsButton =
            "<button id=\"" + RibbonIds.Attachments + "\" label=\"Вложения\" size=\"large\" getImage=\"GetButtonImage\"" +
            " getEnabled=\"OnGetEnabled\"" +
            " screentip=\"Отправить вложения письма\"" +
            " supertip=\"Выбрать отдельные вложения текущего письма и отправить в Telegram\"" +
            " onAction=\"OnSendAttachments\"/>";

        private const string MaxButton =
            "<button id=\"" + RibbonIds.Max + "\" label=\"" + RibbonLabels.Max + "\" size=\"large\" getImage=\"GetButtonImage\"" +
            " getEnabled=\"OnGetEnabled\"" +
            " screentip=\"Отправить в MAX\"" +
            " supertip=\"Сохраняет документ и отправляет его выбранному чату MAX с подписью\"" +
            " onAction=\"OnSendMax\"/>";

        // Печать — стандартный диалог Outlook (Ctrl+P), где можно выбрать «Microsoft Print to PDF».
        private const string PrintButton =
            "<button id=\"" + RibbonIds.Print + "\" label=\"Печать (PDF)\" size=\"large\" getImage=\"GetButtonImage\"" +
            " getEnabled=\"OnGetEnabled\"" +
            " screentip=\"Печать письма (Ctrl+P)\"" +
            " supertip=\"Открывает стандартный диалог печати Outlook — выберите «Microsoft Print to PDF»\"" +
            " onAction=\"OnPrintPdf\"/>";

        public static readonly string Xml = Build( false );
        public static readonly string OutlookXml = Build( true );

        private static string Build( bool withAttachments )
        {
            string attachments = withAttachments ? ( AttachmentsButton + PrintButton ) : "";

            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\"" +
                   " onLoad=\"OnRibbonLoad\">" +
                   "<ribbon>" +
                   "<tabs>" +
                   "<tab id=\"" + RibbonIds.Tab + "\" label=\"Отправить\">" +
                   "<group id=\"TelegramShareGroup\" label=\"Telegram\">" +
                   "<button id=\"" + RibbonIds.SendVia + "\" label=\"Отправить в Telegram\"" +
                   " size=\"large\" getImage=\"GetButtonImage\" getEnabled=\"OnGetEnabled\"" +
                   " screentip=\"Отправить через Telegram\"" +
                   " supertip=\"Сохраняет документ и открывает в Telegram окно «Переслать…» — выберите чат вручную\"" +
                   " onAction=\"OnSendViaTelegram\"/>" +
                   "<button id=\"" + RibbonIds.SendMessage + "\" label=\"" + RibbonLabels.SendMessage + "+сообщение\"" +
                   " size=\"large\" getImage=\"GetButtonImage\" getEnabled=\"OnGetEnabled\"" +
                   " screentip=\"Отправить в Telegram с сообщением\"" +
                   " supertip=\"Сохраняет текущий документ и отправляет его выбранному контакту с подписью\"" +
                   " onAction=\"OnSendButton\"/>" +
                   attachments +
                   "</group>" +
                   "<group id=\"TelegramShareMaxGroup\" label=\"MAX\">" +
                   MaxButton +
                   "</group>" +
                   "<group id=\"TelegramShareLinksGroup\" label=\"Проект\">" +
                   "<button id=\"" + RibbonIds.PinQat + "\" label=\"В заголовок окна\" size=\"large\"" +
                   " getImage=\"GetButtonImage\" screentip=\"Закрепить кнопки в заголовке окна\"" +
                   " supertip=\"Добавляет кнопки «Отправить в Telegram» и «Отправить в MAX» на панель быстрого доступа рядом с кнопкой «Сохранить». Достаточно нажать один раз\"" +
                   " onAction=\"OnPinToQat\"/>" +
                   "<button id=\"" + RibbonIds.Update + "\" label=\"Проверить обновления\" size=\"large\"" +
                   " imageMso=\"Refresh\" screentip=\"Проверить обновления Telegram Share\"" +
                   " supertip=\"Проверяет новую версию и предлагает установить обновление\"" +
                   " onAction=\"OnCheckUpdates\"/>" +
                   "</group>" +
                   "</tab>" +
                   "</tabs>" +
                   "</ribbon>" +
                   "</customUI>";
        }
    }
}
