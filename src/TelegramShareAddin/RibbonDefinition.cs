namespace TelegramShareAddin
{
    /// <summary>Разметка ленты Office, общая для Word и Excel.</summary>
    internal static class RibbonDefinition
    {
        public const string Xml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "<ribbon>" +
            "<tabs>" +
            "<tab id=\"TelegramShareTab\" label=\"Отправить в Telegram\">" +
            "<group id=\"TelegramShareGroup\" label=\"Отправка документа\">" +
            "<button id=\"TelegramShareSendViaButton\" label=\"Отправить в Telegram\"" +
            " size=\"large\" getImage=\"GetButtonImage\" screentip=\"Отправить через Telegram\"" +
            " supertip=\"Сохраняет документ и открывает в Telegram окно «Переслать…» — выберите чат вручную\"" +
            " onAction=\"OnSendViaTelegram\"/>" +
            "<button id=\"TelegramShareButton\" label=\"Отправить в Telegram+сообщение\"" +
            " size=\"large\" getImage=\"GetButtonImage\" screentip=\"Отправить в Telegram с сообщением\"" +
            " supertip=\"Сохраняет текущий документ и отправляет его выбранному контакту с подписью\"" +
            " onAction=\"OnSendButton\"/>" +
            "</group>" +
            "</tab>" +
            "</tabs>" +
            "</ribbon>" +
            "</customUI>";
    }
}