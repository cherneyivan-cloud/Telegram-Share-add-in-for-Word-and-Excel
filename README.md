# Telegram Share — надстройки для Word, Excel и Outlook

[![Скачать](https://img.shields.io/badge/Download-Releases-blue)](https://github.com/cherneyivan-cloud/Telegram-Share-add-in-for-Word-and-Excel/releases)
[![GitHub](https://img.shields.io/badge/GitHub-repository-black)](https://github.com/cherneyivan-cloud/Telegram-Share-add-in-for-Word-and-Excel)

Отправка текущего документа Word, книги Excel или письма Outlook как вложения в **Telegram Desktop** или **MAX** одной кнопкой.

**Скачать:** [Releases](https://github.com/cherneyivan-cloud/Telegram-Share-add-in-for-Word-and-Excel/releases) → `TelegramShareAddin.zip` → распаковать → запустить `install.bat` (без прав администратора).

Работает через локальные приложения Telegram Desktop и MAX (без Bot API(:, через буфер обмена, SendKeys и UI Automation.

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
- Telegram Desktop (64-битный(, установленный и авторизованный — для Telegram
- MAX для Windows, установленный и авторизованный — для MAX
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
- файл попадёт в буфер обмена как файл **вместе с полным путём текстом**;
- откроется Telegram Desktop (при необходимости через протокол `tg://`( или MAX, откроется чат с контактом;
- надстройка подождёт（настройка, по умолчанию 2.5 с**, затем сделает фокус на окно мессенджера、 вставит файл (`Ctrl+V`(, добавит подпись и отправит сообщение.

**Важно**: во время автоматической отправки (несколько секунд** не переключайте окна и не печатайте — клавиши уйдут в активное окно.

После сохранения **полный путь к файлу остаётся в буфере обмена** — его можно вставить в любое текстовое поле (`Ctrl+V`**. В мессенджере `Ctrl+V` по-прежнему вставляет сам файл. Временные файлы автоматически вычищаются из `%TEMP%\TelegramShare` при следующих отправках (старше суток**.




## Кнопки на ленте (Word, Excel, Outlook)

Вкладка называется **«Отправить»**. Кнопки:

- **«Отправить в Telegram»** — спрашивает формат (PDF/DOCX для Word, PDF/XLSX для Excel; в Outlook — MSG без выбора), сохраняет и открывает в Telegram окно **«Переслать…»**; чат выбираете вручную.
- **«Отправить в Telegram+сообщение»** — выбор получателя из списка чатов Telegram (или вручную) и текста сообщения, затем автоматическая отправка.
- **«Отправить в MAX»** — тот же диалог, но с чатами MAX: выбираете чат или вводите username, добавляете сообщение и формат; файл отправляется автоматически.
- **«Вложения»** (только Outlook) — выбор отдельных вложений текущего письма; выбранные файлы отправляются в Telegram (окно «Переслать…», можно несколько сразу).
- **«Печать (PDF)»** (только Outlook) — открывает стандартный диалог печати Outlook (Ctrl+P); выберите «Microsoft Print to PDF», чтобы сохранить письмо в PDF.

В диалоге отправки можно переключить **Мессенджер** — Telegram или MAX — и список чатов подгрузится из выбранного приложения.

Надстройка устанавливается для **Word**, **Excel** и **Outlook** (общий код, три COM-объекта). В Outlook «документ» — это текущее письмо: для Telegram оно сохраняется в **.msg**, а PDF делается кнопкой **«Печать (PDF)»** через стандартный диалог печати (Ctrl+P) и «Microsoft Print to PDF».

Пункт **ПКМ → Отправить → Telegram** работает для любых файлов в Проводнике — `install.ps1` создаёт ярлык `Telegram.exe -sendpath` в папке SendTo.

## Контакты

В диалоге отправки есть **список контактов**: сохранённые вами контакты и **чаты, загруженные прямо из запущенного мессенджера** (их реальный список). Получателя можно выбрать кликом мыши. Кнопка **«Обновить из Telegram»/«Обновить из MAX»** перечитывает список чатов (загружается автоматически при открытии окна). Кнопка **«В список»** добавляет введённый контакт в сохранённые, **«Удалить из списка»** — убирает выделенный.

Над списком есть **поле фильтра** — начните вводить первые буквы, и список сузится. **Групповые чаты и каналы Telegram** показаны **курсивом**.

Если выбран чат из списка, надстройка открывает его **кликом по нему в мессенджере** (а не по ссылке), после чего вставляет файл. Если контакт введён вручную, для Telegram используется `tg://`, а для MAX — ссылка `max://max.ru/<username>` **только при явном `@username`**. Для MAX можно также указать **имя чата как оно отображается в списке** — надстройка найдёт его прокруткой списка (без ссылок, чтобы не открывалась страница «Ссылка устарела»).

Сохранённые контакты хранятся в `%APPDATA%\TelegramShareAddin\config.json` в поле `RecentContacts`.

## Настройки

Файл конфигурации: `%APPDATA%\TelegramShareAddin\config.json`（создаётся автоматически(**:

```json
{
  "RecentContacts" : [  "durov"  ],
  "LastFormat" : "PDF",
  "LastMessenger" : "Telegram",
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
| `LastMessenger` | Последний мессенджер («Telegram» / «MAX»( | Telegram |
| `RememberContact` | Запоминать контакт и формат | true |



## Архитектура и технические детали

- Класс `Addin` реализует `IDTExtensibility2`（жизненный цикл COM-надстройки（ и `IRibbonExtensibility`（лента(**. Интерфейс `IDTExtensibility2` объявлен в коде самостоятельно—— это убирает зависимость от COM-библиотеки «Extensibility»（Extensibility.tlb(.
- Лента задаётся декларативно (XML customUI(;callback кнопки:`TelegramShareAddin.Addin.OnSendButton`（вызов через IDispatch(**.
- Отправка идёт без Bot API:: глубокие ссылки `tg://resolve?domain=...`, `tg://resolve?phone=...`, `tg://user?id=...и** + буфер обмена（`Clipboard.SetFileDropList`( + `SendKeys`（`^v`, подпись, `{ENTER}`(。
- Настройки — Newtonsoft.Json, `%APPDATA%\TelegramShareAddin\config.json`。 Сборка Newtonsoft.Json.dll грузится из папки надстройки через `AssemblyResolve`, поэтому не зависит от версий, уже загруженных хостом.

- Сборка: .NET Framework 4.8, `x64`（обязательно для 64-битного Word(, без strong-name, COM-регистрация пользовательская（в `HKCU`**.



## Устранение неполадок

**Вкладка «Отправить» не появилась**

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



**MAX: «Не удалось открыть чат…»**

- Убедитесь, что MAX для Windows запущен и вы авторизованы;
- нажмите **«Обновить из MAX»** и выберите чат из списка;
- для username используется ссылка `max://max.ru/<username>` — указывайте его явно, через `@` (например, `@support`);
- если чат не находится, откройте его в MAX вручную и повторите отправку.

**Контact не открылся/«контакт не найден»**

- Проверьте точный username（в Telegram имена — латиница和 цифры（;
- для номера телефона начните с `+`（например, `+79123456789`（;
- если `phone`-ссылка не сработала（не для всех клиентов, используйте username или ID（через `@userinfobot`（и числовой ID（.



**Файл очень большой**

Telegram Desktop сам предложит сжатие/подтверждение при отправке больших файлов（лимиты Telegram: до 2 GB на ботах/премиум, 2 GB на десктопе без лимита по размеру для обычных чатов（. Отправка займёт больше времени；надстройка не ждёт завершения загрузки——печатать/нажимать больше не нужно, Telegram сам отправит файл после загрузкии.



**Безопасность**

- Временные файлы лежат в `%TEMP%\TelegramShare` и вычищаются при следующих отправках (старше суток);
- после отправки в буфере обмена остаётся **полный путь к файлу** (текстом) — так задумано;
- токены и боты не используются; надстройка лишь управляет локальными приложениями Telegram и MAX пользователя.

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