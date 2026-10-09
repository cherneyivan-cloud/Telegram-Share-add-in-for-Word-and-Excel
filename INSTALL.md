# Установка Telegram Share для Word и Excel

Надстройки позволяют отправить текущий документ Word или книгу Excel в Telegram Desktop.

## Требования

- Windows 10/11 (x64)
- Microsoft Word и/или Excel **64-бит** (Office 2016/2019/2021/365, включая LTSC)
- .NET Framework 4.8 (входит в состав Windows)
- Telegram Desktop, запущенный и авторизованный

## Быстрая установка (без прав администратора)

1. Распакуйте архив `TelegramShareAddin.zip` в любую папку.
2. Запустите **`install.bat`** двойным щелчком. **Права администратора не требуются.**
3. Перезапустите Word и Excel. На ленте появится вкладка **«Отправить в Telegram»**.
4. Проверка установки: **`check_install.bat`**.
5. Удаление: **`uninstall.bat`** (при желании удалить и файлы: `uninstall.bat -RemoveFiles`).

### Что делает установщик

- копирует файлы в `%APPDATA%\TelegramShareAddin\`;
- регистрирует COM-объекты в `HKCU\Software\Classes\` (Word и Excel), права администратора не нужны;
- добавляет ключи в `HKCU\Software\Microsoft\Office\Word\Addins\TelegramShareAddin`
  и `HKCU\Software\Microsoft\Office\Excel\Addins\TelegramShareExcel` (`LoadBehavior = 3`);
- сбрасывает список сбойных надстроек Office;
- создаёт ярлык **ПКМ → Отправить → Telegram** (ключ `-sendpath`) для любых файлов.

## Использование

Вкладка **«Отправить в Telegram»** на ленте Word/Excel:

- **«Отправить в Telegram»** — спрашивает формат (PDF/XLSX/DOCX), сохраняет документ и открывает в Telegram окно **«Переслать…»** — чат выбираете вручную.
- **«Отправить в Telegram+сообщение»** — выбор получателя из списка чатов Telegram (или вручную) и текст сообщения, затем автоматическая отправка.

В Проводнике для любых файлов: **ПКМ → Отправить → Telegram → Переслать…**

## Сборка из исходников (для разработчиков)

Требования: Visual Studio 2019/2022 (нагрузка «Разработка классических приложений .NET»), установленный Office (PIAs).

1. Откройте `TelegramShareAddin.sln`.
2. Восстановите пакеты NuGet и соберите конфигурацию **Release**.
3. Проекты:
   - `src\TelegramShareAddin` — надстройка для Word (выход: `TelegramShareAddin.dll`);
   - `src\TelegramShareExcelAddin` — надстройка для Excel (выход: `TelegramShareExcel.dll`);
   - общий код (`AddinBase.cs`, `TelegramSender.cs`, `SendForm.cs`, `Settings.cs`, `FormatForm.cs`, `Icons.cs`, `RibbonDefinition.cs`) используется обоими проектами.
4. Сборка переносимого пакета:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-dist.ps1
```

Создаст папку `dist\` и архив `TelegramShareAddin.zip`.

## Публикация на GitHub

Если `git` установлен, выполните в корне проекта:

```powershell
git init
git add .
git commit -m "Telegram Share add-in for Word and Excel"
git branch -M main
git remote add origin https://github.com/<ВАШ_ЛОГИН>/telegram-share-addin.git
git push -u origin main
```

Если установлен GitHub CLI (`gh`) и выполнен вход:

```powershell
gh repo create telegram-share-addin --public --source . --push
```

Замените `<ВАШ_ЛОГИН>` и имя репозитория на свои.

## Устранение неполадок

- **Вкладка не появилась**: Файл → Параметры → Надстройки → Управление «Надстройки COM» → Перейти → включить `TelegramShareAddin` (Word) / `TelegramShareExcel` (Excel).
- **Excel**: на некоторых сборках Office вкладка появляется после явного включения надстройки в диалоге «Надстройки COM» (Excel→Файл→Параметры→Надстройки). Для отправки книги в любом случае работает **ПКМ → Отправить → Telegram** (или кнопки, если вкладка видна).
- **Ошибки при работе**: лог `%APPDATA%\TelegramShareAddin\addin.log`.
- **DOCX/XLSX**: если `SaveCopyAs` недоступен, надстройка копирует сохранённый файл — сохраните документ перед отправкой.
- **Скрипты**: сохраняйте `.ps1` в **UTF-8 с BOM**, иначе Windows PowerShell 5.1 испортит кириллицу.
- **PDF/DOCX/XLSX** экспортируются во временный файл, который удаляется через 10 минут после отправки.
