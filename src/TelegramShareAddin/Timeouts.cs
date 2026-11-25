namespace TelegramShareAddin
{
    /// <summary>
    /// Таймауты работы с мессенджерами. Раньше значения были разбросаны по коду
    /// магическими числами, из-за чего их нельзя было согласованно изменить.
    /// </summary>
    internal static class Timeouts
    {
        /// <summary>Шаг опроса UI Automation.</summary>
        public const int PollStepMs = 300;

        /// <summary>Ожидание главного окна мессенджера.</summary>
        public const int WindowWaitMs = 15000;

        /// <summary>Ожидание появления диалога отправки.</summary>
        public const int SendDialogMs = 8000;

        /// <summary>Поиск чата в списке.</summary>
        public const int ChatOpenMs = 8000;

        /// <summary>Ожидание принятия файла/отправки.</summary>
        public const int AttachmentMs = 12000;

        /// <summary>Короткая пауза после переключения окна.</summary>
        public const int ShortWaitMs = 800;

        /// <summary>Пауза после действия интерфейса.</summary>
        public const int AnimationMs = 600;
    }
}
