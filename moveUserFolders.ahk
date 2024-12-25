#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, Force  ; Запрещаем одновременный запуск нескольких копий скрипта
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2  ; Режим поиска подстроки в заголовке окна

; Переменные
sleepTime := 30000

; Вызов функции для получения диска
GetUserDrive()
return

GetUserDrive() {
    drives := GetAvailableDrives()  ; Получаем список доступных дисков

    ; Проверяем, есть ли доступные диски
    if (drives = "") {
        MsgBox, Нет доступных дисков.
        return
    }

    ; Создаём окно GUI с выбором диска
    Gui, New
    Gui, Add, Text,, Выберите диск для перемещения папок:
    Gui, Add, DropDownList, vSelectedDrive w200 Center, %drives%
    Gui, Add, Button, gButtonMove w100 Center, ОК
    Gui, Show

    return
}

GetAvailableDrives() {
    drives := ""  ; Инициализируем переменную
    Loop, 26 {
        driveLetter := Chr(A_Index + 64) . ":"  ; Получаем букву диска
        if (FileExist(driveLetter))  ; Проверяем, существует ли диск
            drives .= driveLetter . "|"
    }
    StringTrimRight, drives, drives, 1  ; Убираем последний символ "|"
    return drives
}

MigrateFolders() {
    global SelectedDrive  ; Объявляем глобальную переменную

    ; Выполнение PowerShell команды и сохранение результата в клипборд
    RunWait, PowerShell -Command "(Get-WmiObject -Class Win32_ComputerSystem).UserName | ForEach-Object { $_ } | Set-Clipboard", , Hide

    ; Чтение результата из клипборда
    ClipWait  ; Ожидаем, пока данные скопируются в клипборд
    OutputVar := Clipboard

    ; Извлекаем имя пользователя после символа '\'
    StringSplit, parts, OutputVar, \
    userName := parts2

    ; Создаём путь к новой папке
    userFolderPath := SelectedDrive . "\User folders\" . userName . "\"

    ; Открываем Этот компьютер
    Run explorer.exe shell:MyComputerFolder
    WinWaitActive, Этот компьютер

    Sleep, 2000
    Send {Up}
    Sleep, 100
    Send {Down}

    Loop 7  ; Повторяем 7 раз
    {
        Sleep, 100
        Send {AppsKey}
        Sleep, 500
        Send {Up}
        Sleep, 100
        Send {Enter}
        WinWaitActive, Свойства
        Send +{Tab}
        Sleep, 100
        Send {Up}
        Sleep, 100
        Send {Tab}
        Sleep, 100
        Send {End}
        Sleep, 100
        Send ^{Left}
        Sleep, 100
        Send +{Home}
        Sleep, 100
        Send %userFolderPath%
        Sleep, 100
        Send {Enter}
        WinWaitActive, Создание папки 
        Sleep, 100
        Send {Enter}
        WinWaitActive, Переместить папку
        Sleep, 100
        Send {Enter}
        Sleep, 100

        WinActivate, Этот компьютер
        WinWaitActive, Этот компьютер
        Send {Right}
    }

    ExitApp
}

ButtonMove:
    ; Сохраняем выбор пользователя
    Gui, Submit, NoHide  ; Не закрываем окно
    ; Проверяем выбран ли диск
    if (SelectedDrive = "") {
        MsgBox, Пожалуйста, выберите диск перед продолжением.
    } else {
        Gui, Destroy  ; Закрываем окно после успешного выбора
        MigrateFolders()  ; Вызываем функцию для миграции папок
    }
return