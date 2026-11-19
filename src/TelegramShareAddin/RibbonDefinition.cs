namespace TelegramShareAddin
{
    /// <summary>Разметка ленты Office, общая для Word, Excel и Outlook.</summary>
    internal static class RibbonDefinition
    {
        // Кнопка «Вложения» нужна только Outlook — в Word/Excel ленте её нет.
        private const string AttachmentsButton =
            "<button id=\"TelegramShareAttachmentsButton\" label=\"Вложения\" size=\"large\" getImage=\"GetButtonImage\"" +
            " screentip=\"Отправить вложения письма\"" +
            " supertip=\"Выбрать отдельные вложения текущего письма и отправить в Telegram\"" +
            " onAction=\"OnSendAttachments\"/>";

        private const string MaxButton =
            "<button id=\"TelegramShareMaxButton\" label=\"Отправить в MAX\" size=\"large\" getImage=\"GetButtonImage\"" +
            " screentip=\"Отправить в MAX\"" +
            " supertip=\"Сохраняет документ и отправляет его выбранному чату MAX с подписью\"" +
            " onAction=\"OnSendMax\"/>";

        // Печать — стандартный диалог Outlook (Ctrl+P), где можно выбрать «Microsoft Print to PDF».
        private const string PrintButton =
            "<button id=\"TelegramSharePrintButton\" label=\"Печать (PDF)\" size=\"large\" getImage=\"GetButtonImage\"" +
            " screentip=\"Печать письма (Ctrl+P)\"" +
            " supertip=\"Открывает стандартный диалог печати Outlook — выберите «Microsoft Print to PDF»\"" +
            " onAction=\"OnPrintPdf\"/>";

        public static readonly string Xml = Build( false );
        public static readonly string OutlookXml = Build( true );

        private static string Build( bool withAttachments )
        {
            string attachments = withAttachments ? ( AttachmentsButton + PrintButton ) : "";

            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                   "<ribbon>" +
                   "<tabs>" +
                   "<tab id=\"TelegramShareTab\" label=\"Отправить\">" +
                   "<group id=\"TelegramShareGroup\" label=\"Telegram\">" +
                   "<button id=\"TelegramShareSendViaButton\" label=\"Отправить в Telegram\"" +
                   " size=\"large\" getImage=\"GetButtonImage\" screentip=\"Отправить через Telegram\"" +
                   " supertip=\"Сохраняет документ и открывает в Telegram окно «Переслать…» — выберите чат вручную\"" +
                   " onAction=\"OnSendViaTelegram\"/>" +
                   "<button id=\"TelegramShareButton\" label=\"Отправить в Telegram+сообщение\"" +
                   " size=\"large\" getImage=\"GetButtonImage\" screentip=\"Отправить в Telegram с сообщением\"" +
                   " supertip=\"Сохраняет текущий документ и отправляет его выбранному контакту с подписью\"" +
                   " onAction=\"OnSendButton\"/>" +
                   attachments +
                   "</group>" +
                   "<group id=\"TelegramShareMaxGroup\" label=\"MAX\">" +
                   MaxButton +
                   "</group>" +
                   "<group id=\"TelegramShareLinksGroup\" label=\"Проект\">" +
                   "<button id=\"TelegramSharePinQatButton\" label=\"В заголовок окна\" size=\"large\"" +
                   " getImage=\"GetButtonImage\" screentip=\"Закрепить кнопки в заголовке окна\"" +
                   " supertip=\"Добавляет кнопки «Отправить в Telegram» и «Отправить в MAX» на панель быстрого доступа рядом с кнопкой «Сохранить». Достаточно нажать один раз\"" +
                   " onAction=\"OnPinToQat\"/>" +
                   "<button id=\"TelegramShareUpdateButton\" label=\"Проверить обновления\" size=\"large\"" +
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
