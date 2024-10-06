;@Ahk2Exe-SetProductName Windows environment variables editor
;@Ahk2Exe-SetFileVersion 1.0.1.0

;@Ahk2Exe-Set Author, https://github.com/flipeador
;@Ahk2Exe-Set Source, https://gist.github.com/flipeador/df43c2f742585e4599669ced56ea3dda

;@Ahk2Exe-UpdateManifest 0

#Requires AutoHotkey v2.0.18
#SingleInstance Off
#NoTrayIcon

TITLE := 'Environment variables editor'

if WinExist(TITLE . ' ahk_class AutoHotkeyGUI')
    WinActivate(), ExitApp()

USER_KEY := 'HKCU\Environment'
SYSTEM_KEY := 'HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment'

CORE_USER_ENVVARS := [
    'AppData', 'ComputerName', 'HOMEDRIVE', 'HOMEPATH',
    'LocalAppData' , 'UserProfile', 'UserDomain', 'UserName'
]

CORE_SYSTEM_ENVVARS := [
    'AllUsersProfile', 'CommonProgramFiles', 'CommonProgramFiles(x86)',
    'CommonProgramW6432', 'ProgramData', 'ProgramFiles', 'ProgramFiles(x86)',
    'ProgramW6432', 'Public', 'SystemDrive', 'SystemRoot'
]

IL := IL_Create()
IL_Add(IL, 'shell32.dll', -00016) ; scope
IL_Add(IL, 'shell32.dll', -00246) ; name
IL_Add(IL, 'shell32.dll', -00001) ; value
IL_Add(IL, 'shell32.dll', -00003) ; value (file)
IL_Add(IL, 'shell32.dll', -00008) ; value (drive)
IL_Add(IL, 'shell32.dll', -00004) ; value (directory)

UI := Gui('', TITLE)
UI.MarginX := UI.MarginY := 0
UI.SetFont('', 'Consolas')
UI.BackColor := 0x0D1117
TV := UI.AddTreeView(
    Format(
        'w{} R{} c{} Background{} ImageList{}',
        700, 25, '131313', 'EDEDD1', IL
    )
)
SetWindowTheme(TV.Hwnd)
TV.OnEvent('ContextMenu', TV_OnContextMenu)
UI.OnEvent('Close', (*) => ExitApp())

UI.Show()
LoadEnvVars()

#HotIf WinActive(ui)
F5:: Reload()
XButton1:: TV_MoveEnvVarValue('Up')
XButton2:: TV_MoveEnvVarValue('Down')
#HotIf

ShowEditDialog(&result, title, name:='', value:='', flags:=0, id:=0)
{
    UI.Opt('+Disabled')

    dlg := Gui('-MaximizeBox -MinimizeBox +Owner' . UI.Hwnd, title)
    dlg.SetFont('', 'Consolas')
    dlg.AddText('', 'Name:    ')
    dlg.AddComboBox('vName w500 yp R10 Choose1 Limit5', [name])
    .OnEvent('Change', UpdateSaveButton)
    dlg.AddText('xm', 'Value:   ')
    dlg.AddComboBox('vValue w500 yp R10', ParseEnvVarValue(value))
    .OnEvent('Change', OnValueChange)
    dlg['Value'].Text := value
    dlg.AddText('xm', 'Expanded:')
    dlg.AddEdit('vExpanded w500 yp R1 ReadOnly Disabled')
    dlg.AddButton('vSave xm Default', 'Save').OnEvent('Click', OnClose)
    dlg.AddButton('vFile yp', 'Select file').OnEvent('Click', OnSelectFile)
    dlg.AddButton('vDir yp', 'Select directory').OnEvent('Click', OnSelectFile)
    dlg.AddCheckbox('vExpand yp', 'REG_EXPAND_SZ').OnEvent('Click', OnExpandCheck)
    dlg.AddCheckbox('vUnique yp', 'UNIQUE')
    dlg.OnEvent('Close', OnClose), dlg.OnEvent('Escape', OnClose)
    if flags & 0x1 ; REG_EXPAND_SZ?
        dlg['Expanded'].Opt('-Disabled'), dlg['Expand'].Value := true
    if flags & 0x2 ; new/edit value?
        dlg['Name'].Opt('+Disabled'), dlg['Unique'].Value := true
    OnValueChange2()
    UpdateSaveButton()
    UI.GetClientPos(&X, &Y)
    dlg.Show(Format('x{} y{}', X + 10, Y + 10))

    WinWaitClose(dlg)
    UI.Opt('-Disabled')
    WinActivate(UI)
    return result

    OnValueChange(*)
    {
        UpdateSaveButton()
        SetTimer(OnValueChange2, -1000)
    }

    OnValueChange2(*)
    {
        if !dlg
            return
        value := dlg['Value'].Text
        dlg['Expanded'].Text := ExpandEnvVars(title, value)
    }

    OnSelectFile(obj, *)
    {
        dlg.Opt('+OwnDialogs')
        path := ParseEnvVarValue(dlg['Expanded'].Text)
        path := path.Length ? path[1] : ''
        options := obj.Name == 'Dir' ? 'D2' : '3'
        if path := FileSelect(options, path)
            dlg['Value'].Text := path, OnValueChange2()
    }

    UpdateSaveButton(*)
    {
        name := dlg['Name'].Text, value := dlg['Value'].Text
        dlg['Save'].Opt((name == '' || value == '' ? '+' : '-') . 'Disabled')
    }

    OnExpandCheck(cb, *)
    {
        dlg['Expanded'].Opt((cb.Value ? '-' : '+') . 'Disabled')
    }

    OnClose(obj, *)
    {
        if result := obj.Name == 'Save'
        {
            if dlg['Unique'].Value
            {
                values := Map()
                if flags & 0x2 ; new/edit value?
                    for value in GetEnvVarValues(id, true)
                        values.Set(NormalizeEnvVarValue(value, title), 0)
                expanded := ExpandEnvVars(title, dlg['Value'].Text)
                for value in ParseEnvVarValue(expanded)
                {
                    normalized := NormalizeEnvVarValue(value)
                    if values.Has(normalized)
                        return ShowInfo(dlg,
                            'The value contains at least one duplicate.'
                            . '`n`n' . value
                        )
                    values.Set(normalized, 0)
                }
            }
            result := dlg.Submit()
            result.name := ControlGetText(dlg['Name'])
            result.value := ControlGetText(dlg['Value'])
        }
        dlg := dlg.Destroy()
    }
}

NormalizeEnvVarValue(value, scope:=0)
{
    value := ExpandEnvVars(scope, value)
    if SubStr(value, 2, 1) == ':' ; is path?
        value := RegExReplace(value, '[\\/]*$')
    return StrLower(value)
}

TV_GetEnvVarValue()
{
    item := TV.GetSelection()
    items := TV_GetItems(tv, item, 3)
    if items.Length !== 3
    || items[1].text == 'Process'
        return
    return {
        id: item,
        pid: items[2].id,
        scope: StrLower(items[1].text),
        name: items[2].text,
        value: items[3].text
    }
}

TV_SwapEnvVarValues(item, id)
{
    value1 := TV.GetText(item.id)
    value2 := TV.GetText(id)
    icon1 := GetEnvVarValueIcon(item.scope, value1)
    icon2 := GetEnvVarValueIcon(item.scope, value2)
    TV.Modify(id, 'Vis Select Icon' . icon1, value1)
    TV.Modify(item.id, 'Icon' . icon2, value2)
}

TV_MoveEnvVarValue(direction)
{
    if !(item := TV_GetEnvVarValue())
    || !ReadEnvVar(item.scope, item.name, &expand)
    || !(id := direction = 'Up' ? TV.GetPrev(item.id) : TV.GetNext(item.id))
        return
    TV_SwapEnvVarValues(item, id)
    values := GetEnvVarValues(item.pid)
    if !DeleteEnvVar(item.scope, item.name)
    || !WriteEnvVar(item.scope, item.name, values, expand)
        return TV_SwapEnvVarValues(item, id)
    UpdateEnvVars()
}

TV_OnContextMenu(tv, item, rc, x, y)
{
    m := Menu()
    items := TV_GetItems(tv, item, 3)
    isProcess := items.Length >= 1
        && items[1].text = 'Process'
    try env := StrLower(items[1].text)

    Menu_NewVar(*)
    {
        title := 'New ' . env . ' environment variable'
        if !ShowEditDialog(&var, title)
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

    Menu_NewValue(*)
    {
        title := 'New ' . env . ' environment variable value'
        name := items[2].text
        value := ReadEnvVar(env, name, &flags)
        if !ShowEditDialog(&var, title, name, '', flags|2, items[2].id)
            return
        value := ReadEnvVar(env, name, &flags)
        value := RegExReplace(value, '^;+|;+$')
        var.value := StrReplace(var.value, '"', '')
        if InStr(var.value, ';')
            var.value := '"' . var.value . '"'
        value := value . ';' . var.value
        if !DeleteEnvVar(env, name)
        || !WriteEnvVar(env, name, value, var.expand)
            return
        ParseEnvVarValue(var.value, items[2].id, env, 'Vis Select')
        UpdateEnvVars()
    }

    Menu_EditVar(*)
    {
        title := 'Edit ' . env . ' environment variable'
        name := items[2].text
        value := ReadEnvVar(env, name, &expand)
        if !ShowEditDialog(&var, title, name, value, expand)
        || !DeleteEnvVar(env, name)
        || !WriteEnvVar(env, var.name, var.value, var.expand)
            return
        tv.Delete(item)
        id := tv.Add(var.name, items[1].id, 'Vis Select Sort Icon2')
        ParseEnvVarValue(var.value, id, env)
        UpdateEnvVars()
    }

    Menu_DeleteVar(*)
    {
        UI.Opt('+OwnDialogs')
        name := items[2].text
        if MsgBox(
            'Are you sure you want to delete the environment variable?'
            . '`n`n' . name
            ,, 'Icon? YN'
        ) = 'No'
        || !DeleteEnvVar(env, name)
            return
        tv.Delete(item)
        UpdateEnvVars()
    }

    Menu_EditValue(*)
    {
        title := 'Edit ' . env . ' environment variable value'
        name := items[2].text
        value := items[3].text
        ReadEnvVar(env, name, &flags)
        if !ShowEditDialog(&var, title, name, value, flags|2, items[2].id)
        || !DeleteEnvVar(env, name)
            return
        var.value := StrReplace(var.value, '"', '')
        icon := GetEnvVarValueIcon(env, var.value)
        tv.Modify(item, 'Icon' . icon, var.value)
        values := GetEnvVarValues(items[2].id)
        WriteEnvVar(env, name, values, var.expand)
        UpdateEnvVars()
    }

    Menu_DeleteValue(*)
    {
        UI.Opt('+OwnDialogs')
        name := items[2].text
        values := ReadEnvVar(env, name, &expand)
        if MsgBox(
            'Are you sure you want to delete the value?'
            . '`n`n' . items[3].text
            ,, 'Icon? YN'
        ) = 'No'
        || !DeleteEnvVar(env, name)
            return
        tv.Delete(item)
        values := GetEnvVarValues(items[2].id)
        WriteEnvVar(env, name, values, expand)
        UpdateEnvVars()
    }

    tv.Modify(item, 'Select')

    ; ENV
    if !isProcess && items.Length == 1
    {
        m.Add('New variable', Menu_NewVar)
    }

    ; ENV NAME
    if !isProcess && items.Length == 2
    {
        m.Add('New value', Menu_NewValue)
        m.Add()
        m.Add('Edit variable', Menu_EditVar)
        m.Add('Delete variable', Menu_DeleteVar)
    }

    ; ENV NAME VALUE
    if !isProcess && items.Length == 3
    {
        m.Add('Edit value', Menu_EditValue)
        m.Add('Delete value', Menu_DeleteValue)
        m.Add()
        m.Add('Move up', (*) => TV_MoveEnvVarValue('Up'))
        if !tv.GetPrev(item)
            m.Disable('Move up')
        m.Add('Move down', (*) => TV_MoveEnvVarValue('Down'))
        if !tv.GetNext(item)
            m.Disable('Move down')
    }

    if !isProcess && items.Length
        m.Add()
    m.Add('Refresh', LoadEnvVars)

    m.Show()
}

TV_GetItems(tv, item, count)
{
    items := []
    loop item ? count : 0
        items.InsertAt(1, { id: item, text: tv.GetText(item) })
    until !(item := tv.GetParent(item))
    return items
}

LoadEnvVars(*)
{
    TV.Opt('-Redraw')
    TV.Delete()

    user := TV.Add('User')
    ParseUserEnvVars(user)

    system := TV.Add('System')
    ParseSystemEnvVars(system)

    process := TV.Add('Process')
    ParseProcessEnvVars(process)

    TV.Opt('+Redraw')
}

UpdateEnvVars()
{
    ; HWND_BROADCAST | WM_SETTINGCHANGE
    SendMessage(0xFFFF, 0x1A, 0, 'Environment')
    WinRedraw(TV)
}

ParseUserEnvVars(id)
{
    ParseEnvVars(USER_KEY, id)
}

ParseSystemEnvVars(id)
{
    ParseEnvVars(SYSTEM_KEY, id)
}

ParseEnvVars(key, pid)
{
    loop reg key
    {
        if A_LoopRegType != 'REG_SZ' && A_LoopRegType != 'REG_EXPAND_SZ'
            continue

        name := A_LoopRegName
        value := RegRead(key, name)

        id := TV.Add(name, pid, 'Icon2')
        ParseEnvVarValue(value, id, key)
    }

    TV.Modify(pid, 'Bold Sort Expand')
}

ParseEnvVarValue(value, id:=0, scope:=0, options:='')
{
    index := 1
    quote := false
    values := ['']

    loop parse value
    {
        if A_LoopField == '"'
        {
            quote := !quote
            continue
        }
        if !quote && A_LoopField == ';'
            ++index, values.push('')
        else values[index] .= A_LoopField
    }

    values2 := []
    for value in values
        if value !== ''
            values2.Push(value)

    if !id
        return values2

    for value in values2
    {
        TV.Add(
            value, id,
            Format('Icon{} {}',
                GetEnvVarValueIcon(scope, value),
                options
            )
        )
    }
}

GetEnvVarValues(id, arr:=false)
{
    values := arr ? [] : ''
    id := TV.GetChild(id)
    while id
    {
        value := TV.GetText(id)
        if arr
            values.Push(value)
        else
        {
            if InStr(value, ';')
                value := '"' . value . '"'
            values .= (values == '' ? '' : ';') . value
        }
        id := TV.GetNext(id)
    }
    return values
}

GetEnvVarValueIcon(scope, value)
{
    expanded := ExpandEnvVars(scope, value)
    if !InStr(expanded, ':')
        return 3
    isfile := FileExist(expanded), isdir := InStr(isfile, 'D')
    return isdir ? StrLen(expanded) < 4 ? 5 : 6 : isfile ? 4 : 3
}

ParseProcessEnvVars(pid)
{
    for name, value in GetEnvironmentStrings()
    {
        id := TV.Add(name, pid, 'Icon2')
        ParseEnvVarValue(value, id, 0)
    }
    TV.Modify(pid, 'Bold Sort Expand')
}

ExpandEnvVars(scope, str)
{
    return InStr(scope, 'User')
        || InStr(scope, 'HKCU')
        ? ExpandUserEnvVars(str)
        : InStr(scope, 'System')
        || InStr(scope, 'HKLM')
        ? ExpandSystemEnvVars(str)
        : str
}

ExpandUserEnvVars(str)
{
    str := ExpandRegEnvVars(USER_KEY, str)
    str := ExpandCoreEnvVars(CORE_USER_ENVVARS, str)
    return ExpandSystemEnvVars(str)
}

ExpandSystemEnvVars(str)
{
    str := ExpandRegEnvVars(SYSTEM_KEY, str)
    str := ExpandCoreEnvVars(CORE_SYSTEM_ENVVARS, str)
    return str
}

ExpandRegEnvVars(key, str)
{
    str2 := str
    loop reg key
        if A_LoopRegType == 'REG_SZ'
        || A_LoopRegType == 'REG_EXPAND_SZ'
            str2 := StrReplace(str2, '%' . A_LoopRegName . '%', RegRead())
    return str == str2 ? str : ExpandUserEnvVars(str2)
}

ExpandCoreEnvVars(arr, str)
{
    str2 := str
    for name in arr
        str2 := StrReplace(str2, '%' . name . '%', EnvGet(name))
    return str == str2 ? str : ExpandCoreEnvVars(arr, str2)
}

GetEnvironmentStrings()
{
    vars := Map()
    ptr := p := DllCall('GetEnvironmentStringsW', 'Ptr')
    while (str := StrGet(p, 'UTF-16'))
    {
        var := StrSplit(str, '=')
        if var.length == 2
            vars.Set(var[1], var[2])
        p += StrPut(str)
    }
    DllCall('FreeEnvironmentStringsW', 'Ptr', ptr)
    return vars
}

SendMessage(hWnd, Msg, wParam:=0, lParam:=0, Timeout:=1)
{
    hWnd := IsObject(hWnd) ? hWnd.Hwnd : hWnd
    wType := wParam is String ? 'Str' : 'Ptr'
    lType := lParam is String ? 'Str' : 'Ptr'
    r := DllCall('SendMessageTimeoutW', 'Ptr', hWnd, 'UInt', Msg, wType, wParam
        , lType, lParam, 'UInt', 0x22, 'UInt', Timeout, 'PtrP', &result:=0)
    return r ? result : ''
}

SetWindowTheme(hWnd)
{
    DllCall('uxtheme\SetWindowTheme', 'Ptr', hWnd, 'Str', 'Explorer', 'Ptr', 0)
}

ShowInfo(ui, message)
{
    ui.Opt('+OwnDialogs')
    MsgBox(message,, 'Iconi')
}

CreateEnvVar(env, name, value, expand)
{
    UI.Opt('+OwnDialogs')
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
            throw Error('')
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

DeleteEnvVar(env, name)
{
    key := env = 'User' ? USER_KEY : SYSTEM_KEY
    try RegDelete(key, name)
    catch Error as e {
        UI.Opt('+OwnDialogs')
        MsgBox(e.Message,, 'IconX')
        return false
    }
    return true
}

WriteEnvVar(env, name, value, expand)
{
    key := env = 'User' ? USER_KEY : SYSTEM_KEY
    type := expand ? 'REG_EXPAND_SZ' : 'REG_SZ'
    try RegWrite(value, type, key, name)
    catch Error as e {
        UI.Opt('+OwnDialogs')
        MsgBox(e.Message,, 'IconX')
        return false
    }
    return true
}

ReadEnvVar(env, name, &expand)
{
    loop reg env = 'User' ? USER_KEY : SYSTEM_KEY
    {
        if A_LoopRegName = name
        {
            expand := A_LoopRegType == 'REG_EXPAND_SZ' ? 0x1 : 0x0
            return expand || A_LoopRegType == 'REG_SZ' ? RegRead() : ''
        }
    }
}
