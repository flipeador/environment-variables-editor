;@Ahk2Exe-SetFileVersion 1.0.4.0
;@Ahk2Exe-SetProductName Environment Variables Editor
;@Ahk2Exe-SetDescription Environment variables editor
;@Ahk2Exe-SetCopyright https://github.com/flipeador/environment-variables-editor

#Requires AutoHotkey v2.0.19
#SingleInstance Off
#NoTrayIcon

TITLE := 'Environment variables editor'

if WinExist(TITLE . ' ahk_class AutoHotkeyGUI')
    WinActivate(), ExitApp()

USER_KEY := 'HKCU\Environment'
SYSTEM_KEY := 'HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment'

CORE_USER_ENVVARS := [
    'AppData', 'ComputerName', 'HomeDrive', 'HomePath',
    'LocalAppData' , 'UserProfile', 'UserDomain', 'UserName'
]

CORE_SYSTEM_ENVVARS := [
    'AllUsersProfile', 'CommonProgramFiles', 'CommonProgramFiles(x86)',
    'CommonProgramW6432', 'ProgramData', 'ProgramFiles', 'ProgramFiles(x86)',
    'ProgramW6432', 'Public', 'SystemDrive', 'SystemRoot'
]

il := IL_Create()
IL_Add(il, 'shell32.dll', -00016) ; scope
IL_Add(il, 'shell32.dll', -00246) ; name
IL_Add(il, 'shell32.dll', -00001) ; value
IL_Add(il, 'shell32.dll', -00003) ; value (file)
IL_Add(il, 'shell32.dll', -00008) ; value (drive)
IL_Add(il, 'shell32.dll', -00004) ; value (directory)

ui := Gui('+Resize', TITLE)
ui.MarginX := ui.MarginY := 0
ui.SetFont(, 'Consolas')
ui.BackColor := 0x0D1117
tv := ui.AddTreeView(
    Format(
        'w{} R{} c{} Background{} ImageList{}',
        700, 25, '131313', 'EDEDD1', il
    )
)
SetExplorerTheme(tv.Hwnd)
SetBackdropEffect(ui.Hwnd, 'MicaAlt')
tv.OnEvent('ContextMenu', TV_OnContextMenu)
ui.OnEvent('Size', OnSize)
ui.OnEvent('Close', (*) => ExitApp())
ui.Show()

LoadEnvVars()

#HotIf WinActive(ui)
F5:: Reload()
XButton1:: TV_MoveEnvVarValue('Up')
XButton2:: TV_MoveEnvVarValue('Down')
#HotIf

OnSize(*) {
    WinGetClientPos(,, &w, &h, ui.Hwnd)
    ControlMove(0, 0, w, h, tv.Hwnd)
}

ShowEditDialog(title, name:='', value:='', flags:=0, id:=0) {
    ui.Opt('+Disabled')

    dlg := Gui('-MaximizeBox -MinimizeBox +Owner' . ui.Hwnd, title)
    dlg.SetFont(, 'Consolas')
    dlg.AddText(, 'Name:    ')
    dlg.AddComboBox('vName w500 yp R10 Choose1 Limit5', [name])
    dlg['Name'].OnEvent('Change', UpdateSaveButton)
    dlg.AddText('xm', 'Value:   ')
    dlg.AddComboBox('vValue w500 yp R10', ParseEnvVarValue(value))
    dlg['Value'].OnEvent('Change', OnValueChange)
    dlg['Value'].Text := value
    dlg.AddText('xm', 'Expanded:')
    dlg.AddEdit('vExpanded w500 yp R1 ReadOnly Disabled')
    dlg.AddButton('vSave xm Default', 'Save')
    dlg['Save'].OnEvent('Click', OnClose)
    dlg.AddButton('vFile yp', 'Select file')
    dlg['File'].OnEvent('Click', OnSelectFile)
    dlg.AddButton('vDir yp', 'Select directory')
    dlg['Dir'].OnEvent('Click', OnSelectFile)
    dlg.AddCheckbox('vExpand yp', 'REG_EXPAND_SZ')
    dlg['Expand'].OnEvent('Click', OnExpandCheck)
    dlg.AddCheckbox('vUnique yp', 'UNIQUE')
    dlg.AddCheckbox('vStrict yp Checked', 'STRICT')
    dlg['Strict'].OnEvent('Click', UpdateSaveButton)
    dlg.OnEvent('Close', OnClose)
    dlg.OnEvent('Escape', OnClose)
    if flags & 1 ; REG_EXPAND_SZ?
        dlg['Expand'].Value := true
      , dlg['Expanded'].Opt('-Disabled')
    if flags & 2 ; new/edit value?
        dlg['Unique'].Value := true
      , dlg['Name'].Opt('+Disabled')
    OnValueChange2()
    UpdateSaveButton()
    SetBackdropEffect(dlg.Hwnd, 'Acrylic')
    ui.GetClientPos(&X, &Y)
    dlg.Show(Format('x{} y{}', X + 10, Y + 10))

    result := unset
    WinWaitClose(dlg)
    WinActivate(ui)
    return result

    OnValueChange(*) {
        UpdateSaveButton()
        SetTimer(OnValueChange2, -1000)
    }

    OnValueChange2(*) {
        if !dlg
            return
        value := dlg['Value'].Text
        dlg['Expanded'].Text := ExpandEnvVars(title, value)
    }

    OnSelectFile(obj, *) {
        dlg.Opt('+OwnDialogs')
        path := ParseEnvVarValue(dlg['Expanded'].Text)
        path := path.Length ? path[1] : ''
        options := obj.Name = 'Dir' ? 'D2' : '3'
        if path := FileSelect(options, path) {
            dlg['Value'].Text := path
            OnValueChange2()
            UpdateSaveButton()
        }
    }

    UpdateSaveButton(*) {
        name := dlg['Name'].Text
        value := dlg['Value'].Text
        strict := dlg['Strict'].Value
        if name = '' || value = ''
        || (strict && RegExMatch(name, '[^\w()]'))
            dlg['Save'].Opt('+Disabled')
        else
            dlg['Save'].Opt('-Disabled')
    }

    OnExpandCheck(cb, *) {
        dlg['Expanded'].Opt((cb.Value ? '-' : '+') . 'Disabled')
    }

    OnClose(obj, *) {
        if result := obj.Name = 'Save' {
            if dlg['Unique'].Value {
                values := Map()
                if flags & 0x2 ; new/edit value?
                    for value in GetEnvVarValues(id, true)
                        values.Set(NormalizeEnvVarValue(value, title), 0)
                expanded := ExpandEnvVars(title, dlg['Value'].Text)
                for value in ParseEnvVarValue(expanded) {
                    normalized := NormalizeEnvVarValue(value)
                    if values.Has(normalized)
                        return ShowInfo(dlg
                            , 'The value contains at least one duplicate.'
                            . '`n`n' . value)
                    values.Set(normalized, 0)
                }
            }
            result := dlg.Submit()
            result.name := ControlGetText(dlg['Name'])
            result.value := ControlGetText(dlg['Value'])
        }
        ui.Opt('-Disabled')
        dlg := dlg.Destroy()
    }
}

NormalizeEnvVarValue(value, scope:=0) {
    value := ExpandEnvVars(scope, value)
    if SubStr(value, 2, 1) = ':' ; is path?
        value := RegExReplace(value, '[\\/]*$')
    return StrLower(value)
}

TV_GetEnvVarValue() {
    item := tv.GetSelection()
    items := TV_GetItems(tv, item, 3)
    if items.Length != 3
    || items[1].text = 'Process'
        return
    return {
        id: item,
        pid: items[2].id,
        scope: StrLower(items[1].text),
        name: items[2].text,
        value: items[3].text
    }
}

TV_SwapEnvVarValues(item, id) {
    value1 := tv.GetText(item.id)
    value2 := tv.GetText(id)
    icon1 := GetEnvVarValueIcon(item.scope, value1)
    icon2 := GetEnvVarValueIcon(item.scope, value2)
    tv.Modify(id, 'Vis Select Icon' . icon1, value1)
    tv.Modify(item.id, 'Icon' . icon2, value2)
}

TV_MoveEnvVarValue(direction) {
    if !(item := TV_GetEnvVarValue())
    || !ReadEnvVar(item.scope, item.name, &expand)
    || !(id := direction = 'Up' ? tv.GetPrev(item.id) : tv.GetNext(item.id))
        return
    TV_SwapEnvVarValues(item, id)
    values := GetEnvVarValues(item.pid)
    if !DeleteEnvVar(item.scope, item.name)
    || !WriteEnvVar(item.scope, item.name, values, expand)
        return TV_SwapEnvVarValues(item, id)
    UpdateEnvVars()
}

TV_OnContextMenu(tv, item, *) {
    tv.Modify(item, 'Select')
    items := TV_GetItems(tv, item, 3)
    isProcess := items.Length >= 1
        && items[1].text = 'Process'
    try env := StrLower(items[1].text)

    OnNewVariable(*) {
        title := 'New ' . env . ' environment variable'
        var := ShowEditDialog(title)
        if !var
        || !CreateEnvVar(env, var.name, var.value, var.expand)
            return
        id := tv.GetChild(items[1].id)
        while id
            if tv.GetText(id) = var.name
                tv.Delete(id), id := 0
            else id := tv.GetNext(id)
        id := tv.Add(var.name, items[1].id, 'Sort Icon2')
        ParseEnvVarValue(var.value, id, env, 'Vis Select')
        UpdateEnvVars()
    }

    OnNewValue(*) {
        title := 'New ' . env . ' environment variable value'
        name := items[2].text
        value := ReadEnvVar(env, name, &flags)
        var := ShowEditDialog(title, name, '', flags|2, items[2].id)
        if !var || !DeleteEnvVar(env, name)
            return
        var.value := StrReplace(var.value, '"', '')
        icon := GetEnvVarValueIcon(env, var.value)
        n := items.Length == 3 ? items[3].id : '' ; hItemInsertAfter?
        tv.Add(var.value, items[2].id, Format('Vis Select Icon{} {}', icon, n))
        values := GetEnvVarValues(items[2].id)
        WriteEnvVar(env, name, values, var.expand)
        UpdateEnvVars()
    }

    OnEditVariable(*) {
        title := 'Edit ' . env . ' environment variable'
        name := items[2].text
        value := ReadEnvVar(env, name, &expand)
        var := ShowEditDialog(title, name, value, expand)
        if !var
        || !DeleteEnvVar(env, name)
        || !WriteEnvVar(env, var.name, var.value, var.expand)
            return
        tv.Delete(item)
        id := tv.Add(var.name, items[1].id, 'Vis Select Sort Icon2')
        ParseEnvVarValue(var.value, id, env)
        UpdateEnvVars()
    }

    OnDeleteVariable(*) {
        ui.Opt('+OwnDialogs')
        name := items[2].text
        if MsgBox(
            'Are you sure you want to delete the environment variable?'
            . '`n`n' . name
            , 'Delete variable'
            , 'Icon? YN'
        ) = 'No'
        || !DeleteEnvVar(env, name)
            return
        tv.Delete(item)
        UpdateEnvVars()
    }

    OnEditValue(*) {
        title := 'Edit ' . env . ' environment variable value'
        name := items[2].text
        value := items[3].text
        ReadEnvVar(env, name, &flags)
        var := ShowEditDialog(title, name, value, flags|2, items[2].id)
        if !var || !DeleteEnvVar(env, name)
            return
        var.value := StrReplace(var.value, '"', '')
        icon := GetEnvVarValueIcon(env, var.value)
        tv.Modify(item, 'Icon' . icon, var.value)
        values := GetEnvVarValues(items[2].id)
        WriteEnvVar(env, name, values, var.expand)
        UpdateEnvVars()
    }

    OnDeleteValue(*) {
        ui.Opt('+OwnDialogs')
        name := items[2].text
        values := ReadEnvVar(env, name, &expand)
        if MsgBox(
            'Are you sure you want to delete the value?'
            . '`n`n' . items[3].text
            , 'Delete value'
            , 'Icon? YN'
        ) = 'No'
        || !DeleteEnvVar(env, name)
            return
        tv.Delete(item)
        values := GetEnvVarValues(items[2].id)
        WriteEnvVar(env, name, values, expand)
        UpdateEnvVars()
    }

    cm := Menu()

    ; ENV
    if !isProcess && items.Length = 1
        cm.Add('New variable', OnNewVariable)

    ; ENV NAME [VALUE]
    if !isProcess && items.Length > 1 {
        cm.Add('New value', OnNewValue)
        cm.Add()
    }

    ; ENV NAME
    if !isProcess && items.Length = 2 {
        cm.Add('Edit variable', OnEditVariable)
        cm.Add('Delete variable', OnDeleteVariable)
    }

    ; ENV NAME VALUE
    if !isProcess && items.Length = 3 {
        cm.Add('Edit value', OnEditValue)
        cm.Add('Delete value', OnDeleteValue)
        cm.Add()
        cm.Add('Move up', (*) => TV_MoveEnvVarValue('Up'))
        if !tv.GetPrev(item)
            cm.Disable('Move up')
        cm.Add('Move down', (*) => TV_MoveEnvVarValue('Down'))
        if !tv.GetNext(item)
            cm.Disable('Move down')
    }

    if !isProcess && items.Length
        cm.Add()
    cm.Add('Refresh', LoadEnvVars)

    cm.Show()
}

TV_GetItems(tv, item, count) {
    items := []
    loop item ? count : 0
        items.InsertAt(1, { id: item, text: tv.GetText(item) })
    until !(item := tv.GetParent(item))
    return items
}

LoadEnvVars(*) {
    tv.Opt('-Redraw')
    tv.Delete()

    user := tv.Add('User')
    ParseUserEnvVars(user)

    system := tv.Add('System')
    ParseSystemEnvVars(system)

    process := tv.Add('Process')
    ParseProcessEnvVars(process)

    tv.Opt('+Redraw')
}

UpdateEnvVars() {
    ui.Title := '(*) ' . TITLE
    Update() {
        ; HWND_BROADCAST | WM_SETTINGCHANGE
        SendMessage(0xFFFF, 0x1A, 0, 'Environment')
        ui.Title := TITLE
    }
    WinRedraw(tv)
    SetTimer(Update, -500)
}

ParseUserEnvVars(id) {
    ParseEnvVars(USER_KEY, id)
}

ParseSystemEnvVars(id) {
    ParseEnvVars(SYSTEM_KEY, id)
}

ParseEnvVars(key, pid) {
    loop reg key {
        if A_LoopRegType != 'REG_SZ' && A_LoopRegType != 'REG_EXPAND_SZ'
            continue

        name := A_LoopRegName
        value := RegRead(key, name)

        id := tv.Add(name, pid, 'Icon2')
        ParseEnvVarValue(value, id, key)
    }

    tv.Modify(pid, 'Bold Sort Expand')
}

ParseEnvVarValue(value, id:=0, scope:=0, options:='') {
    index := 1
    quote := false
    values := ['']

    loop parse value {
        if A_LoopField = '"' {
            quote := !quote
            continue
        }
        if !quote && A_LoopField = ';'
            ++index, values.push('')
        else values[index] .= A_LoopField
    }

    values2 := []
    for value in values
        if value != ''
            values2.Push(value)

    if !id
        return values2

    for value in values2
        tv.Add(
            value, id,
            Format('Icon{} {}',
                GetEnvVarValueIcon(scope, value),
                options
            )
        )
}

GetEnvVarValues(id, arr:=false) {
    values := arr ? [] : ''
    id := tv.GetChild(id)
    while id {
        value := tv.GetText(id)
        if arr
            values.Push(value)
        else {
            if InStr(value, ';')
                value := Format('"{}"', value)
            values .= (values = '' ? '' : ';') . value
        }
        id := tv.GetNext(id)
    }
    return values
}

GetEnvVarValueIcon(scope, value) {
    expanded := ExpandEnvVars(scope, value)
    if !InStr(expanded, ':')
        return 3
    isfile := FileExist(expanded), isdir := InStr(isfile, 'D')
    return isdir ? StrLen(expanded) < 4 ? 5 : 6 : isfile ? 4 : 3
}

ParseProcessEnvVars(pid) {
    for name, value in GetEnvironmentStrings()
        id := tv.Add(name, pid, 'Icon2')
      , ParseEnvVarValue(value, id, 0)
    tv.Modify(pid, 'Bold Sort Expand')
}

ExpandEnvVars(scope, str) {
    return InStr(scope, 'User')
        || InStr(scope, 'HKCU')
        ? ExpandUserEnvVars(str)
        : InStr(scope, 'System')
        || InStr(scope, 'HKLM')
        ? ExpandSystemEnvVars(str)
        : str
}

ExpandUserEnvVars(str) {
    str := ExpandRegEnvVars(USER_KEY, str)
    str := ExpandCoreEnvVars(CORE_USER_ENVVARS, str)
    return ExpandSystemEnvVars(str)
}

ExpandSystemEnvVars(str) {
    str := ExpandRegEnvVars(SYSTEM_KEY, str)
    str := ExpandCoreEnvVars(CORE_SYSTEM_ENVVARS, str)
    return str
}

ExpandRegEnvVars(key, str) {
    str2 := str
    loop reg key
        if A_LoopRegType = 'REG_SZ'
        || A_LoopRegType = 'REG_EXPAND_SZ'
            str2 := StrReplace(str2, '%' . A_LoopRegName . '%', RegRead())
    return str == str2 ? str : ExpandUserEnvVars(str2)
}

ExpandCoreEnvVars(arr, str) {
    str2 := str
    for name in arr
        str2 := StrReplace(str2, '%' . name . '%', EnvGet(name))
    return str == str2 ? str : ExpandCoreEnvVars(arr, str2)
}

GetEnvironmentStrings() {
    vars := Map()
    ptr := p := DllCall('GetEnvironmentStringsW', 'Ptr')
    while (str := StrGet(p, 'UTF-16')) {
        var := StrSplit(str, '=')
        if var.length = 2
            vars.Set(var[1], var[2])
        p += StrPut(str)
    }
    DllCall('FreeEnvironmentStringsW', 'Ptr', ptr)
    return vars
}

SendMessage(hWnd, Msg, wParam:=0, lParam:=0, Timeout:=1) {
    hWnd := IsObject(hWnd) ? hWnd.Hwnd : hWnd
    wType := wParam is String ? 'Str' : 'Ptr'
    lType := lParam is String ? 'Str' : 'Ptr'
    r := DllCall('SendMessageTimeoutW', 'Ptr', hWnd, 'UInt', Msg, wType, wParam
        , lType, lParam, 'UInt', 0x22, 'UInt', Timeout, 'PtrP', &result:=0)
    return r ? result : ''
}

ShowInfo(ui, message) {
    ui.Opt('+OwnDialogs')
    MsgBox(message,, 'Iconi')
}

CreateEnvVar(env, name, value, expand) {
    ui.Opt('+OwnDialogs')
    key := env = 'User' ? USER_KEY : SYSTEM_KEY
    type := expand ? 'REG_EXPAND_SZ' : 'REG_SZ'
    try {
        RegRead(key, name) ; throws if not exists
        if MsgBox(
            'The environment variable already exists.'
            . '`nDo you want to overwrite it?'
            . '`n`n' . name
            ,, 'IconX YN'
        ) = 'Yes'
            Throw(Error(''))
    }
    catch {
        try {
            try RegDelete(key, name)
            RegWrite(value, type, key, name)
        }
        catch Error as e {
            MsgBox(e.Message,, 'IconX')
            return false
        }
        return true
    }
}

DeleteEnvVar(env, name) {
    key := env = 'User' ? USER_KEY : SYSTEM_KEY
    try RegDelete(key, name)
    catch Error as e {
        ui.Opt('+OwnDialogs')
        MsgBox(e.Message,, 'IconX')
        return false
    }
    return true
}

WriteEnvVar(env, name, value, expand) {
    key := env = 'User' ? USER_KEY : SYSTEM_KEY
    type := expand ? 'REG_EXPAND_SZ' : 'REG_SZ'
    try RegWrite(value, type, key, name)
    catch Error as e {
        ui.Opt('+OwnDialogs')
        MsgBox(e.Message,, 'IconX')
        return false
    }
    return true
}

ReadEnvVar(env, name, &expand) {
    loop reg env = 'User' ? USER_KEY : SYSTEM_KEY {
        if A_LoopRegName = name {
            expand := A_LoopRegType = 'REG_EXPAND_SZ' ? 0x1 : 0x0
            return expand || A_LoopRegType = 'REG_SZ' ? RegRead() : ''
        }
    }
}

SetBackdropEffect(hWnd, type) {
    if (VerCompare(A_OSVersion, '10.0.22621') < 0)
        return ; not supported

    ; Allows the window frame to be drawn in dark mode colors.
    ; All windows default to light mode regardless of the system setting.
    ; DWMWA_USE_IMMERSIVE_DARK_MODE (W11 B22000) DWM_SYSTEMBACKDROP_TYPE BOOL
    DllCall('dwmapi\DwmSetWindowAttribute', 'Ptr', hWnd, 'Int', 20, 'IntP', 1, 'Int', 4)

    ; Set the system-drawn backdrop material of the window.
    ; DWMWA_SYSTEMBACKDROP_TYPE (W11 B22621) DWM_SYSTEMBACKDROP_TYPE DWMSBT_MAINWINDOW
    type := Map('Default', 0, 'None', 1, 'Mica', 2, 'Acrylic', 3, 'MicaAlt', 4)[type]
    DllCall('dwmapi\DwmSetWindowAttribute', 'Ptr', hWnd, 'Int', 38, 'IntP', type, 'Int', 4)
}

SetExplorerTheme(hWnd) {
    DllCall('uxtheme\SetWindowTheme', 'Ptr', hWnd, 'Str', 'Explorer', 'Ptr', 0)
}
