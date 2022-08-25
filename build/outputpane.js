///<reference path="../hm_jsmode_no_global.d.ts"/>
(function () {
    var guid = "{7A0CD246-7F50-446C-B19D-EF2B332A8763}";
    var op_dllobj = null;
    function get_op_dllobj() {
        if (!op_dllobj) {
            op_dllobj = hidemaru.loadDll(hidemaruGlobal.hidemarudir() + "\\HmOutputPane.dll");
        }
        return op_dllobj;
    }
    function _output(msg) {
        op_dllobj = get_op_dllobj();
        if (op_dllobj) {
            msg = msg.replace(/\r\n/g, "\n").replace(/\n/g, "\r\n");
            return op_dllobj.dllFunc.Output(hidemaruGlobal.hidemaruhandle(0), msg);
        }
        return 0;
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
    function _setBaseDir() {
        op_dllobj = get_op_dllobj();
        if (op_dllobj) {
            return op_dllobj.dllFunc.SetBaseDir(hidemaruGlobal.hidemaruhandle(0));
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
        var handle = _getWindowHandle();
        var ret = hidemaruGlobal.sendmessage(handle, 0x111 /*WM_COMMAND*/, command_id, 0);
        return ret;
    }
    if (typeof (module) != 'undefined' && module.exports) {
        module.exports.output = _output;
        module.exports.push = _push;
        module.exports.pop = _pop;
        module.exports.clear = _clear;
        module.exports.getWindowHandle = _getWindowHandle;
        module.exports.setBaseDir = _setBaseDir;
        module.exports.sendMessage = _sendMessage;
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
        OutputPane = {
            output: _output,
            push: _push,
            pop: _pop,
            clear: _clear,
            getWindowHandle: _getWindowHandle,
            setBaseDir: _setBaseDir,
            sendMessage: _sendMessage,
            "guid": guid
        };
    }
})();
