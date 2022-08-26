/*
 * Copyright (C) 2022 Akitsugu Komiyama
 * under the MIT License
 *
 * outputpane v1.0.2
 */

declare var module: { filename: string, directory: string, exports: any };
declare var OutputPane: any;

(function() {
    var guid = "{7A0CD246-7F50-446C-B19D-EF2B332A8763}";

    var op_dllobj: hidemaru.ILoadDllResult = null;

    function get_op_dllobj(): hidemaru.ILoadDllResult {
        if (!op_dllobj) {
            op_dllobj = hidemaru.loadDll(hidemaruGlobal.hidemarudir() + "\\HmOutputPane.dll");
        }

        return op_dllobj;
    }

    function _output(msg: string): number {

        op_dllobj = get_op_dllobj();

        if (op_dllobj) {
            msg = msg.replace(/\r\n/g, "\n").replace(/\n/g, "\r\n");
            return op_dllobj.dllFunc.Output(hidemaruGlobal.hidemaruhandle(0), msg);
        }

        return 0;
    }

    function _push(): number {
        op_dllobj = get_op_dllobj();

        if (op_dllobj) {
            return op_dllobj.dllFunc.Push(hidemaruGlobal.hidemaruhandle(0));
        }

        return 0;
    }

    function _pop(): number {
        op_dllobj = get_op_dllobj();

        if (op_dllobj) {
            return op_dllobj.dllFunc.Push(hidemaruGlobal.hidemaruhandle(0));
        }

        return 0;
    }

    function _clear(): number {
        var handle = _getWindowHandle();
        var ret = hidemaruGlobal.sendmessage(handle, 0x111 /*WM_COMMAND*/, 1009, 0); //1009=クリア
        return ret;
    }

    function _setBaseDir(dirpath: string): number {
        op_dllobj = get_op_dllobj();

        if (op_dllobj) {
            return op_dllobj.dllFunc.SetBaseDir(hidemaruGlobal.hidemaruhandle(0), dirpath);
        }

        return 0;
    }

    var op_windowhandle = null;
    function _getWindowHandle(): number {

        op_dllobj = get_op_dllobj();

        if (op_dllobj) {
            op_windowhandle = op_dllobj.dllFunc.GetWindowHandle(hidemaruGlobal.hidemaruhandle(0));
        }

        return op_windowhandle;
    }

    function _sendMessage(command_id: number): number {
        let handle = _getWindowHandle();
        let ret = hidemaruGlobal.sendmessage(handle, 0x111/*WM_COMMAND*/, command_id, 0);
        return ret;
    }

    var _OutputPane = {
        output : _output,
        push : _push,
        pop : _pop,
        clear : _clear,
        getWindowHandle : _getWindowHandle,
        setBaseDir : _setBaseDir,
        sendMessage: _sendMessage,
    };

    if (typeof (module) != 'undefined' && module.exports) {
        module.exports = _OutputPane;
    } else {
        if (typeof (OutputPane) != 'undefined') {
            if (OutputPane.guid == null || OutputPane.guid != guid) {
                _output("本モジュールとは異なるOutputPaneが、すでに定義されています。\r\n上書きします。\r\n");
            }

            // 一致していたら上書きはしない
            if (OutputPane.guid == guid) {
                return;
            }
        }

        OutputPane = _OutputPane;
        OutputPane.guid = guid;
    }
})();