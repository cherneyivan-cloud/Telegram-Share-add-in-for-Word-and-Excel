# Telegram Share — надстройки для Word и Excel

[![Скачать](https://img.shields.io/badge/Download-Releases-blue)](https://github.com/cherneyivan-cloud/Telegram-Share-add-in-for-Word-and-Excel/releases)
[![GitHub](https://img.shields.io/badge/GitHub-repository-black)](https://github.com/cherneyivan-cloud/Telegram-Share-add-in-for-Word-and-Excel)

Отправка текущего документа Word / книги Excel как вложения в Telegram Desktop одной кнопкой.

**Скачать:** [Releases](https://github.com/cherneyivan-cloud/Telegram-Share-add-in-for-Word-and-Excel/releases) → `TelegramShareAddin.zip` → распаковать → запустить `install.bat` (без прав администратора).

Работает через локальное приложение Telegram Desktop (без Bot API(:, через буфер обмена и SendKeys.

Поддерживает 64-битный Microsoft Word (LTSC MSO 16.0.14332.20296 и новее(.



## Состав проекта

| Файл | Назначение |
| --- | --- |
| `src\TelegramShareAddin\` | Исходный код надстройки (C# / .NET Framework 4.8( |
| `TelegramShareAddin.sln` | Решение Visual Studio |
| `scripts\install.ps1` | Установка без прав администратора |
| `scripts\uninstall.ps1` | Удаление |
| `scripts\check_install.ps1` | Проверка установки |



## Требования

- Windows 10/11, Microsoft Word 64-битный (LTSC/Office 2019/2021/365((
- .NET Framework 4.8 (входит в состав Windows, при необходимости ставится с сайта Microsoft(
- Telegram Desktop (64-битный(, установленный и авторизованный
- Visual Studio 2019 или 2022 (workload «Разработка классических приложений .NET» + доступ к NuGet( для сборки



## Сборка в Visual Studio

1. Откройте `TelegramShareAddin.sln`
2. Дождитесь восстановления пакетов NuGet (`Microsoft.Office.Interop.Word` 15.0.4797.1004 и `Newtonsoft.Json` 13.0.4(**.
   Если восстановление не запустилось автоматически: ПКМ на решение → *Restore NuGet Packages*.
3. Выберите конфигурацию **Release**.
4. Соберите решение (*Build → Build Solution* либо `Ctrl+Shift+B`**.

Результат появится здесь:

- `src\TelegramShareAddin\bin\Release\TelegramShareAddin.dll`
- `src\TelegramShareAddin\bin\Release\Newtonsoft.Json.dll`
- `src\TelegramShareAddin\bin\Release\Microsoft.Office.Interop.Word.dll`
- и сопутствующие файлы。



Сборка из командной строки (опционально(:

```powershell
cd C:\путь\к\решению
dotnet build TelegramShareAddin.sln -c Release
# или, если установлен MSBuild из Visual Studio:
msbuild TelegramShareAddin.sln /p:Configuration=Release /p:Platform="Any CPU" /m
```



## Установка

Перед установкой закройте Microsoft Word (или перезапустите его после установки(.

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1
```

Скрипт:

- найдёт собранную DLL автоматически в `src\TelegramShareAddin\bin\(x64\)\Release\`, либо примите путь через параметр:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1 -SourcePath .\src\TelegramShareAddin\bin\Release
```

- скопирует файлы в `%APPDATA%\TelegramShareAddin\`
- зарегистрирует COM-объект в `HKCU\Software\Classes\CLSID\{AD3B6C2F-4E5F-4A8B-9C1D-3E4F5A6B7C8D}\InprocServer32` и ProgId `TelegramShareAddin`
- создаст ключ `HKCU\Software\Microsoft\Office\Word\Addins\TelegramShareAddin` значение `LoadBehavior = 3` (автозагрузка(

Права администратора не требуются: всё пишется в куст реестра текущего пользователя (HKCU(**. Telegram не затрагивается вообще.



Проверка установки:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check_install.ps1
```



Удаление:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\uninstall.ps1
# или сразу с удалением файлов:
powershell -ExecutionPolicy Bypass -File .\scripts\uninstall.ps1 -RemoveFiles
```



## Использование

1. Откройте документ в Word.
2. На ленте появится вкладка **Telegram Share** → кнопка **Отправить в Telegram**.
3. Укажите контакт:
   - `@username` или `username`（например, `@durov`),
   - номер телефона с `+`（например, `+79123456789`),
   - числовой ID пользователя（например, `123456789`; узнать можно через бота `@userinfobot`(**.
4. При желании добавьте сообщение（оно станет подписью к файлу( и выберите формат:
   - **PDF (рекомендуется** — документ экспортируется в PDF;
   - **DOCX (копия** — создаётся копия файла документа.

5. Нажмите **Отправить**.

Что произойдёт автоматически:

- Word сохранит документ во временный файл в `%TEMP%`（.pdf/.docx(;
- файл попадёт в буфер обмена как файл;
- откроется Telegram Desktop (при необходимости через протокол `tg://`(, откроется чат с контактом;
- надстройка подождёт（настройка, по умолчанию 2.5 с**, затем сделает фокус на окно Telegram、 вставит файл (`Ctrl+V`(, наберёт подпись和 нажмёт `Enter`。

**Важно**: во время автоматической отправки (несколько секунд** не переключайте окна и не печатайте — клавиши уйдут в активное окно. В конце буфер обмена восстанавливается, временный файл удаляется.**



## Кнопки на ленте (Word и Excel)

Вкладка называется **«Отправить в Telegram»**. Кнопки:

- **«Отправить в Telegram»** — спрашивает формат (PDF/DOCX или XLSX), сохраняет документ и открывает в Telegram окно **«Переслать…»**; чат выбираете вручную.
- **«Отправить в Telegram+сообщение»** — выбор получателя из списка чатов Telegram (или вручную) и текста сообщения, затем автоматическая отправка.

Надстройка устанавливается и для **Word**, и для **Excel** (общий код, два COM-объекта).

Пункт **ПКМ → Отправить → Telegram** работает для любых файлов в Проводнике — `install.ps1` создаёт ярлык `Telegram.exe -sendpath` в папке SendTo.

## Контакты

В диалоге отправки есть **список контактов**: сохранённые вами контакты и **чаты, загруженные прямо из запущенного Telegram Desktop** (их реальный список). Получателя можно выбрать кликом мыши. Кнопка **«Обновить из Telegram»** перечитывает список чатов (загружается автоматически при открытии окна). Кнопка **«В список»** добавляет введённый контакт в сохранённые, **«Удалить из списка»** — убирает выделенный.

Над списком есть **поле фильтра** — начните вводить первые буквы, и список сузится. **Групповые чаты и каналы** показаны **курсивом**.

Если выбран чат из Telegram, надстройка открывает его **кликом по нему в Telegram** (а не по `tg://`-ссылке), после чего вставляет файл. Если контакт введён вручную (`@username`, `+номер`, ID), используется `tg://`.

Сохранённые контакты хранятся в `%APPDATA%\TelegramShareAddin\config.json` в поле `RecentContacts`.

## Настройки

Файл конфигурации: `%APPDATA%\TelegramShareAddin\config.json`（создаётся автоматически(**:

```json
{
  "RecentContacts" : [  "durov"  ],
  "LastFormat" : "PDF",
  "ChatOpenDelayMs" : 2500,
  "PasteDelayMs" : 1200,
  "RememberContact" : true
}
```

| Параметр | Описание | По умолчанию |
| --- | --- | --- |
| `ChatOpenDelayMs` | Задержка после открытия чата перед вставкой файла, мс | 2500 |
| `PasteDelayMs` | Пауза после вставки файла перед набором подписи, мс | 1200 |
| `RecentContacts` | История контактов для автодополнения | |
| `LastFormat` | Последний формат («PDF» / «DOCX»( | PDF |
| `RememberContact` | Запоминать контакт и формат | true |



## Архитектура и технические детали

- Класс `Addin` реализует `IDTExtensibility2`（жизненный цикл COM-надстройки（ и `IRibbonExtensibility`（лента(**. Интерфейс `IDTExtensibility2` объявлен в коде самостоятельно—— это убирает зависимость от COM-библиотеки «Extensibility»（Extensibility.tlb(.
- Лента задаётся декларативно (XML customUI(;callback кнопки:`TelegramShareAddin.Addin.OnSendButton`（вызов через IDispatch(**.
- Отправка идёт без Bot API:: глубокие ссылки `tg://resolve?domain=...`, `tg://resolve?phone=...`, `tg://user?id=...и** + буфер обмена（`Clipboard.SetFileDropList`( + `SendKeys`（`^v`, подпись, `{ENTER}`(。
- Настройки — Newtonsoft.Json, `%APPDATA%\TelegramShareAddin\config.json`。 Сборка Newtonsoft.Json.dll грузится из папки надстройки через `AssemblyResolve`, поэтому не зависит от версий, уже загруженных хостом.

- Сборка: .NET Framework 4.8, `x64`（обязательно для 64-битного Word(, без strong-name, COM-регистрация пользовательская（в `HKCU`**.



## Устранение неполадок

**Вкладка «Telegram Share» не появилась**

1. Проверьте установку: `scripts\check_install.ps1`.
2. Убедитесь, что надстройка включена в Word: *Файл → Параметры → Надстройки → Перейти… → Надстройки COM* → галочка у `TelegramShareAddin`.
3. Перезапустите Word полностью（и, если нужно, ПК（.
4. Если Word 32-битный — проект не загрузится;требуется 64-битная версия Office（см. требования（.



**Ошибка: «Не удалось открыть Telegram»**

- Убедитесь, что Telegram Desktop установлен и выполнен вход;
- убедитесь, что `tg://` ссылки открывают Telegram（Параметры → Система → обработчик протоколов（.

иетки「или попробуйте запустить из PowerShell: `Start-Process "tg://resolve?domain=durov"`.



**Ошибка: «Не найдено окно Telegram Desktop»**

- Откройте главное окно Telegram（оно должно быть развёрнуто, а не свёрнуто в трей（ и повторите;
- увеличьте `ChatOpenDelayMs` в конфиге（например, 4000–6000（.



**Контact не открылся/«контакт не найден»**

- Проверьте точный username（в Telegram имена — латиница和 цифры（;
- для номера телефона начните с `+`（например, `+79123456789`（;
- если `phone`-ссылка не сработала（не для всех клиентов, используйте username или ID（через `@userinfobot`（и числовой ID（.



**Файл очень большой**

Telegram Desktop сам предложит сжатие/подтверждение при отправке больших файлов（лимиты Telegram: до 2 GB на ботах/премиум, 2 GB на десктопе без лимита по размеру для обычных чатов（. Отправка займёт больше времени；надстройка не ждёт завершения загрузки——печатать/нажимать больше не нужно, Telegram сам отправит файл после загрузкии.



**Безопасность**

- Временные файлы удаляются сразу после отправки;
- буфер обмена восстанавливается к исходному содержимому;
- токены и боты не используются; надстройка лишь управляет локальным приложением Telegram пользователя.

## Что не поддерживается

- Telegram Bot API（сознательно, по ТЗ（;
- отправка нескольких файлов сразу;
- scheduling/очередь—— всё делается синхронно при нажатии кнопки（несколько секунд блокировки формы（.



## Лицензия

Проект предоставляется «как есть», без гарантий. Используйте на свой риск.

## Примечания по сборке и запуску скриптов

- Скрипты запускаются **только через `powershell`**, а не как отдельная команда:
  правильно — `powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1`,
  неправильно — `-ExecutionPolicy Bypass -File .\scripts\install.ps1` (будет ошибка «имя не распознано»).
- Скрипт сам находит сборку: `bin\x64\Release`, `bin\Release`, `bin\x64\Debug`, `bin\Debug`.
  Если вы собрали в **Debug**, ничего указывать не нужно.
- Все `.ps1` и `.cs` сохранены в **UTF-8 с BOM**. Если правите их в редакторе, сохраняйте с BOM,
  иначе Windows PowerShell 5.1 прочитает кириллицу неверно и скрипт не запустится.
- COM-объект регистрируется в HKCU как управляемый: `InprocServer32 = mscoree.dll`,
  а также `Class`, `Assembly`, `RuntimeVersion=v4.0.30319`, `CodeBase=file:///...TelegramShareAddin.dll`.
  Права администратора не требуются.
- После установки обязательно **перезапустите Word**. При этом откройте Word и проверьте,
  что надстройка включена: Файл → Параметры → Надстройки → Перейти → Надстройки COM.
- Проект ссылается на настоящий интерфейс `IDTExtensibility2` из `Extensibility.dll`
  (`C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\PublicAssemblies\`).
  Для другой редакции VS (Community/Enterprise) или другого пути поправьте `HintPath` в `.csproj`.
  Использовать самодельное объявление этого интерфейса нельзя: у настоящего метода имеют
  DispId 1–5, иначе Word создаёт объект, но не вызывает `OnConnection` и помечает надстройку сбойной.
- Надстройка ведёт диагностический лог `%APPDATA%\TelegramShareAddin\addin.log`
  (загрузка, `OnConnection`, `GetCustomUI`, ошибки). Полезно при разборе проблем.