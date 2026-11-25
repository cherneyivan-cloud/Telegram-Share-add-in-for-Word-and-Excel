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
        public static readonly string Xml = Build( false );
        public static readonly string OutlookXml = Build( true );

        private static string Build( bool withAttachments )
        {
            string attachments = withAttachments
                ? Button( RibbonIds.Attachments , "Вложения" , "OnSendAttachments" ,
                          "Отправить вложения письма" ,
                          "Выбрать отдельные вложения текущего письма и отправить в Telegram" ) +
                  Button( RibbonIds.Print , "Печать (PDF)" , "OnPrintPdf" ,
                          "Печать письма (Ctrl+P)" ,
                          "Открывает стандартный диалог печати Outlook — выберите «Microsoft Print to PDF»" )
                : string.Empty;

            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                   "<ribbon><tabs>" +
                   $"<tab id=\"{RibbonIds.Tab}\" label=\"Отправить\">" +
                   "<group id=\"TelegramShareGroup\" label=\"Telegram\">" +
                   Button( RibbonIds.SendVia , "Отправить в Telegram" , "OnSendViaTelegram" ,
                           "Отправить в текущий чат Telegram" ,
                           "Сохраняет документ и открывает окно выбора чата Telegram" ) +
                   Button( RibbonIds.SendMessage , RibbonLabels.SendMessage + "+сообщение" , "OnSendButton" ,
                           "Отправить в Telegram с сообщением" ,
                           "Сохраняет текущий документ и отправляет его выбранному контакту с подписью" ) +
                   attachments +
                   "</group>" +
                   "<group id=\"TelegramShareMaxGroup\" label=\"MAX\">" +
                   Button( RibbonIds.Max , RibbonLabels.Max , "OnSendMax" ,
                           "Отправить в MAX" ,
                           "Сохраняет документ и отправляет его выбранному чату MAX с подписью" ) +
                   "</group>" +
                   "<group id=\"TelegramShareLinksGroup\" label=\"Проект\">" +
                   Button( RibbonIds.PinQat , "В заголовок окна" , "OnPinToQat" ,
                           "Закрепить кнопки в заголовке окна" ,
                           "Добавляет кнопки «Отправить в Telegram» и «Отправить в MAX» на панель быстрого доступа рядом с кнопкой «Сохранить». Достаточно нажать один раз" ) +
                   Button( RibbonIds.Update , "Проверить обновления" , "OnCheckUpdates" ,
                           "Проверить обновления Telegram Share" ,
                           "Проверяет новую версию и предлагает установить обновление" ,
                           null , "Refresh" ) +
                   "</group></tab></tabs></ribbon></customUI>";
        }

        /// <summary>Кнопка ленты с единым набором атрибутов.</summary>
        private static string Button( string id , string label , string action , string screentip , string supertip ,
                                      string image = "GetButtonImage" , string imageMso = null )
        {
            string imageAttribute = imageMso != null ? $"imageMso=\"{imageMso}\"" : $"getImage=\"{image}\"";

            return $"<button id=\"{id}\" label=\"{label}\" size=\"large\" {imageAttribute}" +
                   $" screentip=\"{screentip}\" supertip=\"{supertip}\" onAction=\"{action}\"/>";
        }
    }
}
