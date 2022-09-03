/**
 * @file 秀丸のjsmode用のTypeScript定義ファイル
 * @author Akitsugu Komiyama
 * @license MIT
 * @version v1.0.6
 */

declare namespace OutputPane {
    /**
     *
     * アウトプット枠への文字列の出力するメソッドとなります。
     * 
     * @param msg
     * アウトプット枠に表示したい文字列を指定します。    
     * 改行は、「\n」でも「\r\n」でもどちらでも大丈夫です。    
     * 
     * 文字列である必要はなく、数値やオブジェクトにも対応しています。    
     *
     * @returns
     * 成功したら0以外を返す、失敗したら0を返す。
     */
    function output(msg: any): number;

    /**
     *
     * アウトプット枠への文字列の出力するメソッドとなります。    
     * 最後に改行が自動で付与されます。
     * 
     * @param msg
     * アウトプット枠に表示したい文字列を指定します。    
     * 改行は、「\n」でも「\r\n」でもどちらでも大丈夫です。    
     * 
     * 文字列である必要はなく、数値やオブジェクトにも対応しています。    
     *
     * @returns
     * 成功したら0以外を返す、失敗したら0を返す。
     */
     function outputline(msg: any): number;

    /**
     * アウトプット枠内の文字列を一時退避するメソッドとなります。    
     * 後で、pop()メソッドで復元する際に利用します。
     * 
     * @see pop
     *
     * @returns
     * 成功したら0以外を返す、失敗したら0を返す。
     */
    function push(): number;

    /**
     * アウトプット枠内の文字列を一時退避したものを復元するメソッドとなります。    
     * 後で、復元する際に利用します。
     * 
     * @see push
     *
     * @returns
     * 成功したら0以外を返す、失敗したら0を返す。
     */
    function pop(): number;

    /**
     * アウトプット枠内の文字列をクリアするメソッドとなります。    
     *
     * @returns
     * 返り値に意味はありません。
     */
    function clear(): number;    

    /**
     * アウトプット枠内でファイルやディレクトリの相対パスを表示する際の、    
     * 基本となるディレクトリを設定します。    
     * 主にファイルジャンプに影響します。
     * 
     * @param dirpath
     * ベースとして設定したいディレクトリを指定します。
     *
     * @returns
     * 成功したら0以外を返す、失敗したら0を返す。
     */
    function setBaseDir(dirpath: string): number

    /**
     * アウトプット枠内へとコマンドを実行するメソッドとなります。
     *
     * @param command_id
     * アウトプット枠に対して行う、コマンドの番号を指定します。    
     * 
     * アウトプット枠のコマンド値一覧    
     * - 1001 枠を閉じる
     * - 1002 中断
     * - 1005 検索
     * - 1006 次の結果
     * - 1007 前の結果
     * - 1008 タグジャンプ
     * - 1009 クリア
     * - 1010 下候補
     * - 1011 上候補
     * - 1013 すべてコピー
     * - 1014 レジストリ変更を元に色を更新
     * - 1015 先頭にカーソル移動
     * - 1016 最後にカーソル移動
     * - 1100 位置：左
     * - 1101 位置：右
     * - 1102 位置：上
     * - 1103 位置：下
     * 
     * @returns
     * 成功したら0以外を返す、失敗したら0を返す。
     */
    function sendMessage(command_id: number): number

    /**
     * アウトプット枠のハンドルを常に取得することが出来ます。    
     * アウトプット枠に対してネイティブ関数を利用した操作をする際に    
     * このハンドルが必要となります。
     *
     * @returns
     * 成功したら0以外を返す、失敗したら0を返す。
     */
    function getWindowHandle(): number
}