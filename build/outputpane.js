/*
 * Copyright (C) 2022 Akitsugu Komiyama
 * under the MIT License
 *
 * outputpane v1.1.0
 */
/// <reference path="../../hm_jsmode_ts_difinition/types/hm_jsmode_strict.d.ts" />
(function () {
    var guid = "{7A0CD246-7F50-446C-B19D-EF2B332A8763}";
    var op_dllobj = hidemaru.loadDll("HmOutputPane.dll");
    var selfdir = null;
    var hidemaruhandlezero = hidemaru.getCurrentWindowHandle();
    // execjsで読み込まれていたら、{filename,directory}のそれぞれのプロパティに有効な値が入る
    function get_including_by_execjs() {
        var cjf = hidemaruGlobal.currentjsfilename();
        var cmf = hidemaruGlobal.currentmacrofilename();
        if (cjf != cmf) {
            var dir = cjf.replace(/[\/\\][^\/\\]+?$/, "");
            return {
                __filename: cjf,
                __dirname: dir
            };
        }
        return {};
    }
    var selfinfo = get_including_by_execjs();
    if (typeof (selfinfo.__dirname) != 'undefined') {
        selfdir = selfinfo.__dirname;
    }
    else if (typeof (module) != 'undefined' && module.exports) {
        selfdir = __dirname;
    }
    else {
        _output("outputpane.js モジュールが想定されていない読み込み方法で利用されています。\r\n");
        return;
    }
    var op_com = hidemaru.createObject(selfdir + "\\" + "outputpane.dll", "OutputPane.OutputPane");
    if (!op_com) {
        _output("outputpane.dllが読み込めませんでした。\r\n");
        return;
    }
    // 関数の時に、文字列に治す
    function replacer(key, value) {
        if (typeof value === "function") {
            return "[fn]:" + value.toString();
        }
        return value;
    }
    function _stringify(obj, space) {
        if (space === void 0) { space = 2; }
        var text = "";
        if (typeof (obj) == "undefined") { // typeofで判定する
            return undefined;
        }
        text = JSON.stringify(obj, replacer, space);
        if (text) {
            text = text.replace(/\r\n/g, "\n").replace(/\n/g, "\r\n");
        }
        return text;
    }
    function _output(msg) {
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
        return op_dllobj.dllFunc.Output(hidemaruhandlezero, modify_msg);
    }
    function _outputLine(msg) {
        var ret = _output(msg);
        _output("\n");
        return ret;
    }
    function _push() {
        return op_dllobj.dllFunc.Push(hidemaruhandlezero);
    }
    function _pop() {
        return op_dllobj.dllFunc.Push(hidemaruhandlezero);
    }
    function _clear() {
        return _sendMessage(1009);
    }
    function _setBaseDir(dirpath) {
        if (typeof (dirpath) != "string") {
            return 0;
        }
        return op_dllobj.dllFunc.SetBaseDir(hidemaruhandlezero, dirpath);
    }
    var op_windowhandle = null;
    function _getWindowHandle() {
        op_windowhandle = op_dllobj.dllFunc.GetWindowHandle(hidemaruhandlezero);
        return op_windowhandle;
    }
    function _sendMessage(command_id) {
        if (typeof (command_id) != "number") {
            return 0;
        }
        if (op_com) {
            var ret = op_com.OutputPane_SendMessage(command_id);
            return ret;
        }
        return 0;
    }
    var _OutputPane = {
        output: _output,
        outputLine: _outputLine,
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
