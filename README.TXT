readme.txt    
＜参照＞　LS32.txt    '

            update: 2002.8.10 s.f roz(0),roz(1)を突当成形のﾊﾟﾗﾒｰﾀへ'
'            update: 2002.10.16 KYOCERA ﾒﾆｭｰ画面起動時の原点信号出力ON→OFF
'　　　　　　　　　　　　　　　　　　　　　"原点"完了後にOrgON追加
'            update: 2002.10.17 KYOCERA 原点復帰後に初回原点復帰完了ﾌﾗｸﾞgOrgStartFlgをON
'                                       原点信号をﾀｲﾏｰで監視
'                                       原点でないと自動成形ﾓｰﾄﾞ移行不可
'            update: 2002.10.18 KYOCERA 原点表示の修正 If gOrgStartFlg = False Then...End If追加
'            update: 2002.10.25 s.f. Ver．表示修正
'            update: 2002.10.26 s.f. 「真空到達」無効へ
'            update: 2003. 8.26 s.f. * 指定圧力＋２００Ｋｇ以上で非常停止 *
'            update: 2003. 9.11 s.f. LS21_TC　成形終了時の非常停止エラー対策
'            update: 2003. 9.12 s.f. genten()　原点出し後　HiSpeedを指定値に戻す。
'
'            update: 2003.12.15 s.f. LS-32立上げに伴う変更　MplDef.bas　のみ　新規　2003.11.04付け
'　　　　　　　　　　　　　　　　　　これに伴い　PGM_MenuのVERﾅﾝﾊﾞｰを　LS-32　へ変更
'
'            update: 2004. 3. 8 s.f. LS21_SC 変更　成形軸制御モード　’７’追加　（上軸衝突判定付）
'                                    RecEmgDTsave 非常停止メッセージの保存
'


2004.3.8 updata   ・2個打ち防止プログラム
　　　　　　　　　　　軸制御モード「7」新規追加
                      モード7　：　突き当て成形モード「1」と同じ動作＋指定加重の70%の加重が３秒以上
                                   かかったとき非常停止が入る。
　　　　　　　　　・非常停止時のアラーム内容ファイルをディスクに、
　　　　　　　　　　保存する機能追加。
                     新規必須ファイル：EmgMsg.txt　dataフォルダ内に必要
2004.3.20 update  ・LS32で変更した内容を　LS31へ移植
　　　　　　　　　　注意：MplDef.bas　　が　LS31用　: 2002. 1.13付
                                     　　 　LS32用  : 2003.11.04付
2004.3.20 update: s.f. MYEdit.frm　の　SetData(),GetData()　を変更（3/8変更のバグ修正　'edit'の読み込み書き出しエラー）
'　　　　　　　　  　　軸制御ｺﾏﾝﾄﾞ　7追加： 現在有効コマンド 0,1,2,3,7,8,9
'　　
2004.3.30　update:　s.f　3/8変更の部分の　「非常停止　内容表示」のバグ修正

2004. 4.23  update: 2004. 4.23 s.f   timeupで非常停止
2004.4.24'  update: 2004. 4.24 s.f.  LS21_TC内のカウンタ、ﾀｸﾄﾀｲﾑ、表示　改造
'
2004.4.25   update: 2004.4.25  s.f   Myedit　の  VScroll1(j).min = 210 * lK1     "200"を"210"へ変更
2004.5.5    update: 2004.5. 5  s.f   温度係数、肉厚補正ルーチン　追加  PGM_KTD,My_lib,MYEDIT, LS21_SC, LS21_TC,　タイムアップで非常停止「連続成形」にも適用。



