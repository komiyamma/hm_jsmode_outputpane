/*
 * Copyright (C) 2022 Akitsugu Komiyama
 * under the MIT License
 *
 * outputpane v1.0.5
 */

declare var module: { filename: string, directory: string, exports: any };
declare var OutputPane: any;

(function () {
    var guid = "{7A0CD246-7F50-446C-B19D-EF2B332A8763}";

    var op_dllobj: hidemaru.ILoadDllResult = null;
    var hidemaruhandlezero = hidemaruGlobal.hidemaruhandle(0); // このタイミング必須
    function get_op_dllobj(): hidemaru.ILoadDllResult {
        if (!op_dllobj) {
            op_dllobj = hidemaru.loadDll("HmOutputPane.dll");
        }

        return op_dllobj;
    }

    // 関数の時に、文字列に治す
    function replacer(key: string, value: any) {
        if (typeof value === "function") {
            return value.toString();
        }
        return value;
    }

    function _stringify(obj: any, space: number | string = 2): string | undefined {
        var text: string = "";
        if (typeof (obj) == "undefined") { // typeofで判定する
            return undefined;
        }
        var text = JSON.stringify(obj, replacer, space);

        if (text) {
            text = text.replace(/\r\n/g, "\n").replace(/\n/g, "\r\n");
        }

        return text;
    }

    function _output(msg: any): number {

        op_dllobj = get_op_dllobj();

        if (op_dllobj) {
            var modify_msg: string = "";
            if (typeof (msg) == "undefined") {
                modify_msg = "(undefined)";
            } else if (msg == null) {
                modify_msg = "(null)";
            } else if (typeof (msg) == "string") {
                modify_msg = msg
            } else if (typeof (msg) == "object") {
                try {
                    modify_msg = _stringify(msg, 2);
                } catch(e) {
                    modify_msg = msg.toString();
                }
            } else {
                try {
                    modify_msg = _stringify(msg, 2);
                } catch(e) {
                    modify_msg = msg.toString();
                }
            }

            modify_msg = modify_msg.replace(/\r\n/g, "\n").replace(/\n/g, "\r\n");
            return op_dllobj.dllFunc.Output(hidemaruhandlezero, modify_msg);
        }

        return 0;
    }

    function _outputline(msg: any) {
        var ret = _output(msg);
        _output("\n");
        return ret;
    }

    function _push(): number {
        op_dllobj = get_op_dllobj();

        if (op_dllobj) {
            return op_dllobj.dllFunc.Push(hidemaruhandlezero);
        }

        return 0;
    }

    function _pop(): number {
        op_dllobj = get_op_dllobj();

        if (op_dllobj) {
            return op_dllobj.dllFunc.Push(hidemaruhandlezero);
        }

        return 0;
    }

    function _clear(): number {
        var handle = _getWindowHandle();
        var ret = hidemaruGlobal.sendmessage(handle, 0x111 /*WM_COMMAND*/, 1009, 0); //1009=クリア
        return ret;
    }

    function _setBaseDir(dirpath: string): number {
        if (typeof(dirpath) != "string") {
            return 0;
        }

        op_dllobj = get_op_dllobj();

        if (op_dllobj) {
            return op_dllobj.dllFunc.SetBaseDir(hidemaruhandlezero, dirpath);
        }

        return 0;
    }

    var op_windowhandle = null;
    function _getWindowHandle(): number {

        op_dllobj = get_op_dllobj();

        if (op_dllobj) {
            op_windowhandle = op_dllobj.dllFunc.GetWindowHandle(hidemaruhandlezero);
        }

        return op_windowhandle;
    }

    function _sendMessage(command_id: number): number {
        if (typeof(command_id) != "number") {
            return 0;
        }

        let handle = _getWindowHandle();
        let ret = hidemaruGlobal.sendmessage(handle, 0x111/*WM_COMMAND*/, command_id, 0);
        return ret;
    }

    var _OutputPane = {
        output: _output,
        outputline: _outputline,
        push: _push,
        pop: _pop,
        clear: _clear,
        getWindowHandle: _getWindowHandle,
        setBaseDir: _setBaseDir,
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