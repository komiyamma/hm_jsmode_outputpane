/*
 * Copyright (C) 2022 Akitsugu Komiyama
 * under the MIT License
 *
 * outputpane v1.0.4
 */
(function () {
    var guid = "{7A0CD246-7F50-446C-B19D-EF2B332A8763}";
    var op_dllobj = null;
    function get_op_dllobj() {
        if (!op_dllobj) {
            op_dllobj = hidemaru.loadDll(hidemaruGlobal.hidemarudir() + "\\HmOutputPane.dll");
        }
        return op_dllobj;
    }
    // 関数の時に、文字列に治す
    function replacer(key, value) {
        if (typeof value === "function") {
            return value.toString();
        }
        return value;
    }
    function _stringify(obj, space) {
        if (space === void 0) { space = 2; }
        var text = "";
        if (typeof (obj) == "undefined") { // typeofで判定する
            return undefined;
        }
        var text = JSON.stringify(obj, replacer, space);
        if (text) {
            text = text.replace(/\r\n/g, "\n").replace(/\n/g, "\r\n");
        }
        return text;
    }
    function _output(msg) {
        op_dllobj = get_op_dllobj();
        if (op_dllobj) {
            var modify_msg = "";
            if (typeof (msg) == "undefined") {
                modify_msg = "(undefined)";
            }
            else if (msg == null) {
                modify_msg = "(null)";
            }
            else if (typeof (msg) == "string") {
                modify_msg = msg;
            }
            else if (typeof (msg) == "object") {
                try {
                    modify_msg = _stringify(msg, 2);
                }
                catch (e) {
                    modify_msg = msg.toString();
                }
            }
            else {
                try {
                    modify_msg = _stringify(msg, 2);
                }
                catch (e) {
                    modify_msg = msg.toString();
                }
            }
            modify_msg = modify_msg.replace(/\r\n/g, "\n").replace(/\n/g, "\r\n");
            return op_dllobj.dllFunc.Output(hidemaruGlobal.hidemaruhandle(0), modify_msg);
        }
        return 0;
    }
    function _outputline(msg) {
        var ret = _output(msg);
        _output("\n");
        return ret;
    }
    function _push() {
        op_dllobj = get_op_dllobj();
        if (op_dllobj) {
            return op_dllobj.dllFunc.Push(hidemaruGlobal.hidemaruhandle(0));
        }
        return 0;
    }
    function _pop() {
        op_dllobj = get_op_dllobj();
        if (op_dllobj) {
            return op_dllobj.dllFunc.Push(hidemaruGlobal.hidemaruhandle(0));
        }
        return 0;
    }
    function _clear() {
        var handle = _getWindowHandle();
        var ret = hidemaruGlobal.sendmessage(handle, 0x111 /*WM_COMMAND*/, 1009, 0); //1009=クリア
        return ret;
    }
    function _setBaseDir(dirpath) {
        if (typeof (dirpath) != "string") {
            return 0;
        }
        op_dllobj = get_op_dllobj();
        if (op_dllobj) {
            return op_dllobj.dllFunc.SetBaseDir(hidemaruGlobal.hidemaruhandle(0), dirpath);
        }
        return 0;
    }
    var op_windowhandle = null;
    function _getWindowHandle() {
        op_dllobj = get_op_dllobj();
        if (op_dllobj) {
            op_windowhandle = op_dllobj.dllFunc.GetWindowHandle(hidemaruGlobal.hidemaruhandle(0));
        }
        return op_windowhandle;
    }
    function _sendMessage(command_id) {
        if (typeof (command_id) != "number") {
            return 0;
        }
        var handle = _getWindowHandle();
        var ret = hidemaruGlobal.sendmessage(handle, 0x111 /*WM_COMMAND*/, command_id, 0);
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
    }
    else {
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
