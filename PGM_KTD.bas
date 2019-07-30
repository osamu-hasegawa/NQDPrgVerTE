Attribute VB_Name = "PGM_KTD"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    PGM_KTD
'
'         update: 2002.6.29  s.f   difftime
'         update: 2002.10.5  s.f   difftime!
'         update: 2002.12.03 s.f   RecDtsave0, RecDtsave 追加
'         update: 2002.12.07 s.f   RecDtsave0(icnt) へ変更
'         update: 2002.12.09 s.f   cooloff, heatoff 初期リセット　追加
'         update: 2004. 3. 8 s.f   RecEmgDtsave 非常停止メッセージの保存  2004.3.8'
'         update: 2004. 3.12 s.f   速度指令電圧　Global 宣言
'         update: 2004. 3.30 s.f   非常停止ﾒｯｾｰｼﾞバグ修正
'         update: 2004. 5. 5 s.f   温度係数、肉厚補正ルーチン　追加  PGM_KTD,My_lib,MYEDIT, LS21_SC, LS21_TC
'         update: 2005. 9.27 s.f   保温停止モード　追加
'         update: 2005. 9.28 s.f   T係数　表示色変更
'         update: 2005.11. 6 s.f   オーバーフロー対策 idc65536,idc256,ddc05
'         update: 2006.04.14 s.f   on error goto
'         update: 2006.04.15 s.f   error 表示
'         update: 2006.05.15 s.f   data書き込み、読み込み時　”L"　追加
'       Ver.3.33R_070927 2007.09.27 s.f  Z補正　指定したｾｸﾞﾒﾝﾄNo.へ　できるようにする
'       Ver.3.33R_071113 2007.11.13 s.f  「強制ソーク」復活、　「1回成形」enable=Falseへ
'       Ver.3.33R_071119 2007.11.19 s.f  加圧時間制御　バグ修正（edit時、データ継承）、平均値AND最新値で　更新判定へ
'　　　　　　　　　　　　　　　　　　　　加圧時間制御　ON時は、　T係数データの　ファイルからの読み込みしない
'       Ver.3.33R_071120 2007.11.20 s.f  バグ修正、　空成形-排出　追加、　連続成形再開　追加
'       Ver.3.33R_071203 2007.12.03 s.f  非常停止　メッセージ追加、変更
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Ver.NQD_70_080320 2008. 03.20 s.f  C870v1向け　ReadyWaitを　すべてに入れる
'       Ver.NQD_71_080910 2008. 9.10 s.f  新QD　２号機立上に伴う　モーターパルス、回転方向見直し
'       Ver.NQD_71_090817 s.f  NQD6立上に伴う見直し
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
'
Global iFlg_hijyou%                    ' 非常停止FLG　'090817　NQD70_SCから移籍
'                                        0: 正常 1:ﾀｲﾑｱｯﾌﾟ    2:軸制御7     3:4:5:DCエラー
'                                        6: r_pres　1ton越え  7:    8:     9:
'                                       10:　ｼｰｹﾝｻｰ
'
Global InitDat!(0 To 50)               '保存データ
Global InitStr$(0 To 50)
'
Global TPass!(0 To 12000)                '経過時間(秒)
Global ZAxis!(0 To 12000)                '座標（Z-軸）
Global Press!(0 To 12000)                '型締圧
Global Templ!(0 To 12000)                '型温度
Global Templu!(0 To 12000)               '型温度 上
Global Templd!(0 To 12000)               '型温度 下
Global Const ResDtSize = 12000
Global BrdFlg$
Global StartTime!                       'Debug用
Global GCnt0%                           '成形中データカウンタ
Global GCnt1%
Global Const H24Hr = 24# * 3600!        'Timer用 一日の秒数
Global EmgFlg%                          '非常停止
Global gOrgFlg%                         '原点復帰完了=TRUE
'
Global Err_ic%, Err_id%                 'ERROR
Global pv_ch!        '/* マニュアル時の速度／位置切り換え値*/
Global VccLw!                           '真空Zero
Global VccHi!                           '真空到達点
'
Global FrmMenuFlg%                      'メニューから抜けるときfalse
Global gM_sa!                           'メニューの速度制御の時の/* 設定圧力 */
Global gM_p!                           'メニューの速度制御の時の/* 設定Ｐ値 */
Global gM_lim!                           'メニューの速度制御の時の/* 設定リミット値 */
Global ViewFlg%                         '画面番号
'Global Const gDirect! = -1            'S.Mの回転方向 (+1 or -1)  '08.3.24 tsubaki
Global Const gDirect! = 1            'S.Mの回転方向 (+1 or -1)  '08.09.10東栄
Global Const gRev2Disp As Double = 24000       '1回転あたりのパルス数 東栄サーボモーター　レゾルバ 24000Pulse/rev.  昔のLS機
'Global Const gRev2Disp As Double = 16000         '1回転あたりのパルス数 KD2002ED(ﾂﾊﾞｷ)モーター  2000*4=8000
Global gTimeUpCnt%                      'タイムアップのカウンタ
Global gVumFlg%                         '真空到達=1
Global Const idc16777216 As Long = 16777216  '　オーバーフロー対策で追加  2005.11.22
Global Const idc8388607 As Long = 8388607  '　オーバーフロー対策で追加  2005.11.22
Global Const idc65536  As Long = 65536  '　オーバーフロー対策で追加　この下３行　2005.11.6　ｓ．ｆ
Global Const idc256 As Long = 256
Global Const ddc05 As Double = 0.5
Global Const dc0 As Double = 0#
Global Const LongTime As Long = 80000     '80000秒（1日=86400秒）を　異常に長い時間とする。　100116
Global sdt1$, sdt2$, sdt3$             'エラー表示用  2006.4.14
' -----------------  2005.5.追加
Global versionNo$, ppos$      '　version　No.　PGM_Menuの　Label1(13), PresentPosition(現在位置）
Global CmndColoff!(0 To 9)  'コマンド釦の色 off普段の色
Global CmndColon!(0 To 9)  'コマンド釦の色　on　押されたときの色
Global T_keisuCol!(0 To 4)  '温度係数、肉厚補正表示のbackColor
Global kkno%               ' 加圧時間制御　型No
'--------------- レンズ成形機プログラム
' 2001年3月
'
Global gcoxFlName$       'コントロールデータファイル名
Global gcoxFldir$        'ディレクトリ
'
Global gCoxFlDtMax%
Global gCoxDlDt(0 To 200) As String       'coxファイルの読んだままのデータ
Global scom(0 To 200) As String       '
Global sisub(0 To 200) As Long        '
Global sjsub(0 To 200) As Long        '
Global sksub(0 To 200) As Long        '
Global slsub(0 To 200) As Long        '
Global hcomm(0 To 3) As String        '
Global dcomm(0 To 200) As String      '
Global seg_num(0 To 100) As Integer   'セグメント番号
Global ic(0 To 100) As Integer        '制御方式
Global pres(0 To 100) As Single      'プレス圧力
Global z(0 To 100) As Single          '目標位置
Global vel(0 To 100) As Single        '速度
Global t0(0 To 100) As Single         'Time Out
Global p(0 To 100) As Single          'PID P
Global ptime%                         '測定時間 分
Global ytemp%                         '予備加熱 度
'
' ----------   温度係数、肉厚補正データ     2004.4.30

Global ishu%                      ' 成形　？周目
Global ishu_bkup%                 ' 成形　？周目のbackup
Global T_keisuCont%(0 To 3)        '温度係数コントロール
Global T_keisu!(0 To 9)          '温度係数データ
Global T_keisu_dum!
Global Z3_HoseiCont%(0 To 2)          '温度係数コントロール
Global Z3_Hosei!(0 To 9)          '温度係数データ
Global DkatJ!(0 To 1)           '加圧時間目標値
Global AkatJ!(0 To 1)           '加圧時間アラーム設定　上限下限
Global Acp!(0 To 1)             'Cp位置アラーム設定　上限下限
Global Almdisp$                     'アラーム表示　加圧時間ｏｒＣｐ位置
Global AlmON%                   ' alarm On/Off flg
Global iSeikeiTorF_flg%         ' 成形　有効ｏｒ無効ＦＬＧ
Global katCflag%                '加圧時間 自動制御フラグ
Global katDflag%                '加圧時間　制御　変更ダメＦＬＧ
Global Henkou_No%               '加圧時間自動制御用 型変更内容　変更なし、減らす、増やす、入れ替え
Global kaatsuJ!(0 To 10, 0 To 5)
Global iflgKataTorF%(0 To 9)     ' 型フラグ　true=本物　　False=dammy
Global iPltMax%                  '1回成形時の型（パレット）数
Global Saikaiflg As Boolean     ' 再開フラグ2007.11.19 tsuika
Global lSokuFlg%                  '強制ソークタイム
Global Karauchiflg As Boolean   ' 空成形－排出　フラグ
'
Global gDate$                         '結果グラフ日付け
Global gTime$                         '結果グラフ時間
Global gGphDtNum%                     '結果グラフデータ数
Global gResFlName$                    '結果データファイル名
Global gResFldir$                     'ディレクトリ
Global FlNmRecDt$                     '成形データファイル名
Global Rec_of_Mold$                   '成形データ　文字変数
'
Global gErrMsg$(0 To 1, 0 To 36)      'エラーメセージ
Global gemgmsg                        'エラーメセージ
'
Global kataNo$(0 To 10)                 ' 型のナンバー　　　'2007.11.12　tsuika
Global kataNoHyj$(0 To 36)                    ' 型Ｎｏ．　表示用リングバッファ
Global kataNoPnt As Integer                     '　型No.　ポインター
Global katamax                          ' 成形機内の　ステーション数
'--------------- [QD61]LS21_S.C で定義してある変数
Global atemp!(0 To 12000, 0 To 2)    '1801 →12000　へ変更　130425
Global aposi!(0 To 12000)   '1801 →12000　へ変更　130425
Global apre!(0 To 12000)   '1801 →12000　へ変更　130425
Global roz!(2)               '　突当成形ﾊﾟﾗﾒｰﾀ　幅,時間
Global ivd%, id_0%, id_1%, id_2%
'--------------- 手動の位置制御速度設定用
Global gHiSpeed!                      '手動の位置制御速度
Global gLwSpeed!                      '手動の位置制御速度
Global r_pres_kousei!                  '圧力読み取り値のゼロ校正

Global nout%, v!                      'cal_pidの　変数　速度指令電圧

Global gOrgIL As Boolean              '原点インターロック
Global gOrgStartFlg As Boolean        '初回原点復帰完了フラグ
Global wTm0!, wTm1!                   '経過時間計算用     2004.5.12 追加  "ｵｰﾊﾞｰﾌﾛｰ" 対策
'  -----        2009.8.17 追加
Global fintime!, Tm2f1!, Tm2f2!
Public Sub Main()
Dim i%
' On Error GoTo errHandler:
  CmndColoff(1) = &H8000000F     '終了コマンド釦の色　　　　　灰
  CmndColoff(3) = &HC0FFC0       'Vエディトのコマンド釦の色　　薄緑
  CmndColoff(9) = &HC0C0FF       '保温停止のコマンド釦の色　　ピンク
  CmndColoff(0) = &HFFFFC0       '5分停止のコマンド釦の色　　水色
  CmndColon(1) = vbRed '&HFF&    'コマンド釦の色 onのとき　赤　　　　赤
  CmndColon(3) = &HC0FFC0         'コマンド釦の色 onのとき　薄緑
  CmndColon(9) = &HC0C0FF         'コマンド釦の色 onのとき　ピンク
  CmndColon(0) = vbBlue        'コマンド釦の色 onのとき　ao
  T_keisuCol!(0) = &HFFFFFF    '温度係数、肉厚補正　表示backcolor　off 灰色
  T_keisuCol!(1) = &HFFFFC0    '温度係数、肉厚補正　表示backcolor　on　水色
  T_keisuCol!(2) = &H800012    '温度係数、肉厚補正　表示forecolor　on　黒
  T_keisuCol!(3) = &HFF00FF    '温度係数、肉厚補正　表示forecolor　on point中　　ピンク
  T_keisuCol!(4) = &HE0E0E0    '温度係数、肉厚補正　表示backcolor　dummy  水色
  lSokuFlg = False        '強制ソークタイム   通常時は　OFF
  katCflag = False      ' プログラム開始時は、必ず加圧制御OFF
  Karauchiflg = False      ' プログラム開始時は、一旦false
  Saikaiflg = False         'プログラム開始時は、一旦false
  katamax = 6           ' STATION SSU = 6
  ishu = 1              ' 1週目から　スタート

'
    For i = 0 To 9
        kataNo(i) = Format(i + 1, "##")     ' 型Ｎｏ．の初期化
    Next i
    kataNo(10) = " 0"    ' 型Ｎｏ．調整数の初期化＝0
'    ﾀﾞﾐｰ型指定の　reset
  For i = 0 To 9
    iflgKataTorF(i) = True
  Next i
'
ppos = "KTD"
'
  InitDtLoad
  cfileLoad
  coxDtRead gcoxFldir & gcoxFlName
  coxDtSet
  BoardInit
  ResetOFF          '/* リセット　ｏｆｆ */
  SetErrMsg         'アラームメッセージ
  'DebugData         'Debug
  gResFlName = "*.mpr"                  '結果データファイル名
  gResFldir = App.path & "\..\data\"  'ディレクトリ
  'ADMain.Show
  InitStr(2) = "roz.con"                    'ロボットデータファイル名
  InitStr(3) = App.path & "\..\robo\"       'ディレクトリ
  'IOChk.Show '
  ViewFlg = 1
  gOrgFlg = False                       '原点復帰完了=TRUE
  gTimeUpCnt = 0                    'タイムアップのカウンタ
  gVumFlg = 0                       '真空到達=1
  
'  VacuumOFF                        '2006.1221 削除　s.f.
  SeikeiOFF                         '2006.12.21 新規 s.f
'
  CoolOFF
  HeatOFF
'
  ReadyFrm.Show
  'PGM_Menu.Show
  Exit Sub
'
''
End Sub

Public Sub coxFlLoad()
Dim fDir$, fname$, rflg%
    
    fname = gcoxFlName        'コントロールデータファイル名
    fDir = gcoxFldir          'ディレクトリ
    rflg = False
    Call GenFile.SetCtrl("ファイル読込", "読込", "取消")
    Call GenFile.SetFile(cLoad, fDir, fname, "*.cox")
    GenFile.Show vbModal
    Call GenFile.GetFile(rflg, fDir, fname)
    Set GenFile = Nothing
    If rflg Then
      Screen.MousePointer = 11
      '
      coxDtRead fDir$ & fname
      gcoxFlName = fname      'コントロールデータファイル名
      gcoxFldir = fDir        'ディレクトリ
      '
      Screen.MousePointer = 0
    End If
End Sub

Public Sub coxDtRead(fl$)
'　080312:　NQD対応
Dim i%, fnum%, l%
Dim dmy$, dt$, com$, dta$(0 To 4)
Dim iaf%, ja%
Dim isub As Long
Dim jsub As Long
Dim ksub As Long
Dim lsub As Long

  fnum = FreeFile
  Open fl For Input As #fnum
    For l = 0 To 7
      Line Input #fnum, gCoxDlDt(l)
    Next l
    '
    For l = 0 To 2: hcomm(l) = gCoxDlDt(l): Next l
    l = 4: ptime = Val(gCoxDlDt(l))      '測定時間
    l = 6: ytemp = Val(gCoxDlDt(l))      '予備加熱温度
    l = 7
    '軸駆動制御コマンドの読込
    For i = 0 To 100
      Line Input #fnum, dt
      l = l + 1
      gCoxDlDt(l) = dt
      seg_num(i) = Val(Mid(dt, 1, 2))
      ic(i) = Val(Mid(dt, 4, 4))
      z(i) = Val(Mid(dt, 9, 9))
      vel(i) = Val(Mid(dt, 19, 10))
      pres(i) = Val(Mid(dt, 30, 8))
      t0(i) = Val(Mid(dt, 39, 8))
      p(i) = Val(Mid(dt, 48, 6))
      If ic(i) = 9 Then Exit For
    Next i
    'データを読み取る
    Input #fnum, dmy
    l = l + 1
    gCoxDlDt(l) = dmy
    ja = 0
    For i = 0 To 200
      Line Input #fnum, dt
      l = l + 1
      gCoxDlDt(l) = dt
      scom(i) = Mid(dt, 1, 2)
      isub = Val(Mid(dt, 4, 5))
      com = Left(scom(i), 1)
      Select Case com
      Case "S", "L"                  ' 2006.5.15 "L" 追加 s.f
        iaf = iaf + 1
        jsub = Val(Mid(dt, 10, 5))
        ksub = Val(Mid(dt, 16, 5))
        lsub = Val(Mid(dt, 22, 5))
      Case "J"
        iaf = iaf + 1
      Case "H"                      ' 2007.11.12 "H" 追加 s.f
        iaf = iaf + 1
      Case "P"
        ja = ja + 1
        If Right(scom(i), 1) = "R" And isub = 1 And ic(ja - 1) <> 2 Then iaf = iaf + 1
        If Right(scom(i), 1) = "W" And isub = 4 And ic(ja - 1) <> 2 Then iaf = iaf + 1
      Case "E"
        Exit For
      End Select
      sisub(i) = isub
      sjsub(i) = jsub
      sksub(i) = ksub
      slsub(i) = lsub
    Next i
'  -- 温度係数、肉厚補正データ
    Input #fnum, dmy
    l = l + 1
    gCoxDlDt(l) = dmy
    Input #fnum, T_keisuCont(0), T_keisuCont(1)
    l = l + 1
    gCoxDlDt(l) = "  " & Format(T_keisuCont(0), "0.000") & ",  " & Format(T_keisuCont(1), "0.000")
    If (katCflag = True) Then               ' katcflag（加圧制御フラグ）がTrueなら　T係数データをファイルから読み込まない
            Line Input #fnum, dmy           ' １行読み飛ばし
        Else
            Input #fnum, T_keisu(0), T_keisu(1), T_keisu(2), T_keisu(3), T_keisu(4)
    End If
    l = l + 1
    dt = "  " & Format(T_keisu(0), "0.000")
    For i = 1 To 4
      dt = dt & ",  " & Format(T_keisu(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
'
    If (katCflag = True) Then
            Line Input #fnum, dmy                 ' katcflag（加圧制御フラグ）がTrueなら　T係数データをファイルから読み込まない
        Else
            Input #fnum, T_keisu(5), T_keisu(6), T_keisu(7), T_keisu(8), T_keisu(9)
    End If
    l = l + 1
    dt = "  " & Format(T_keisu(5), "0.000")
    For i = 6 To 9
      dt = dt & ",  " & Format(T_keisu(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
 '
    Input #fnum, dmy
    l = l + 1
    gCoxDlDt(l) = dmy
    Input #fnum, Z3_HoseiCont(0), Z3_HoseiCont(1), Z3_HoseiCont(2)
    l = l + 1
    gCoxDlDt(l) = "  " & Format(Z3_HoseiCont(0), "0.000") & ",  " & Format(Z3_HoseiCont(1), "0.000") & ",  " & Format(Z3_HoseiCont(2), "0.000")
    Input #fnum, Z3_Hosei(0), Z3_Hosei(1), Z3_Hosei(2), Z3_Hosei(3), Z3_Hosei(4)
    l = l + 1
    dt = "  " & Format(Z3_Hosei(0), "0.000")
    For i = 1 To 4
      dt = dt & ",  " & Format(Z3_Hosei(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
'
    Input #fnum, Z3_Hosei(5), Z3_Hosei(6), Z3_Hosei(7), Z3_Hosei(8), Z3_Hosei(9)
    l = l + 1
    dt = "  " & Format(Z3_Hosei(5), "0.000")
    For i = 6 To 9
      dt = dt & ",  " & Format(Z3_Hosei(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
'
 '
    Input #fnum, dmy                  '  加圧時間制御ｍａｘ、　ｍｉｎの読み込み
    l = l + 1
    gCoxDlDt(l) = dmy
    Input #fnum, DkatJ(1), DkatJ(0)
    l = l + 1
    gCoxDlDt(l) = "  " & Format(DkatJ(1), "000.0") & ",  " & Format(DkatJ(0), "000.0")
'
    Input #fnum, dmy                  '  型No.　データ　読み込み
    l = l + 1
    gCoxDlDt(l) = dmy
    Input #fnum, kataNo(0), kataNo(1), kataNo(2), kataNo(3), kataNo(4), kataNo(5), kataNo(6), kataNo(7), kataNo(8)
    l = l + 1
    dt = "  " & kataNo(0)
    For i = 1 To 8
        dt = dt + "  " & kataNo(i)
    Next i
    gCoxDlDt(l) = dt
 '
    Input #fnum, dmy                  '  加圧時間&Cp値　ALARM ｍａｘ、　ｍｉｎの読み込み
    l = l + 1
    gCoxDlDt(l) = dmy
'
    Input #fnum, AkatJ(1), AkatJ(0)
    l = l + 1
    gCoxDlDt(l) = "  " & Format(AkatJ(1), "0") & ",  " & Format(AkatJ(0), "0")
'
    Input #fnum, Acp(1), Acp(0)
    l = l + 1
    gCoxDlDt(l) = "  " & Format(Acp(1), "0.000") & ",  " & Format(Acp(0), "0.000")
'
'
  Close fnum
  gCoxFlDtMax = l
  gGphDtMax = iaf       'データ数 元はiaf
End Sub

Public Sub InitDtLoad()
Dim i%, fnum%
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = App.path & "\..\data\"
  flNm = "PGM.ini"
  Open fDir & flNm For Input As #fnum
  For i = 0 To 50
    Input #fnum, InitDat(i), InitStr(i)
  Next i
  Close #fnum
  'gcoxFlName = InitStr(0)       'コントロールデータファイル名
  'gcoxFldir = InitStr(1)        'ディレクトリ
  'InitDat(10)=成形カウンタ
  'InitDat(11)=成形カウンタトウタル
End Sub
Public Sub InitDtSave()
Dim i%, fnum%
Dim fDir$, flNm$
  InitStr(0) = gcoxFlName    'コントロールデータファイル名
  InitStr(1) = gcoxFldir     'ディレクトリ
  fnum = FreeFile
  fDir = App.path & "\..\Data\"
  flNm = "PGM.ini"
  Open fDir & flNm For Output As #fnum
  For i = 0 To 50
    Write #fnum, InitDat(i), InitStr(i)
  Next i
  Close #fnum
End Sub
Public Sub RecDtSave0(icnt!)                     '成形データファイルの作成
Dim j%, fnum%, sdt$
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = App.path & "\..\data\"
'  FlNmRecDt = "LS" & Mid(Date, 6, 2) & Mid(Date, 9, 2) & Mid(Time, 1, 2) & Mid(Time, 4, 2) & Format(Int(icnt), "0") & ".lsl"
  FlNmRecDt = "LS" & Format$(Now, "yymmddhhmmss") & Format(Int(icnt), "0") & ".lsl"
  sdt = " No.     Z3         ct1    ct2"
  sdt = sdt & "      cc1     cc2    cc3"
  sdt = sdt & "    cc3-2     cp         ﾀｸﾄ     T係数    Z3補正"
  Open fDir & FlNmRecDt For Output As #fnum
     Write #fnum, gcoxFlName & "   " & Date$ & "   " & Time$
     Write #fnum, sdt
  Close #fnum
End Sub
Public Sub RecDtSave999()            '成形データのセーブの終了処理　　コントロールデータを追加　　2009.9.12追加
Dim j%, fnum%, l%
Dim fDir$
  fnum = FreeFile
  fDir = App.path & "\..\data\"
  Open fDir & FlNmRecDt For Append As #fnum
    For l = 0 To gCoxFlDtMax
     Write #fnum, gCoxDlDt(l)
    Next l
  Close #fnum
End Sub
Public Sub RecDtSave(Rec_of_Mold$)            '成形データのセーブ
Dim j%, fnum%
Dim fDir$
  fnum = FreeFile
  fDir = App.path & "\..\data\"
  Open fDir & FlNmRecDt For Append As #fnum
     Write #fnum, Rec_of_Mold & "   " & Time$
  Close #fnum
End Sub
Public Sub RecEmgDtSave(sdt1$, sdt2$, sdt3$)            '非常停止データのセーブ  2004.3.8 追加　s.f
Dim j%, fnum%
Dim fDir$, emgmsg$, flNm$
  fnum = FreeFile
  fDir = App.path & "\..\data\"
  flNm = "emgmsg.txt"
     emgmsg = ArmEmgMsgChk$()
  Open fDir & flNm For Append As #fnum
     Write #fnum, Date$ & " " & Time$ & "  " & emgmsg$
     Write #fnum, "  " & sdt1
     Write #fnum, "  " & sdt2
     Write #fnum, "  " & sdt3 & ppos
  Close #fnum
End Sub

Public Sub ResDtSave(i_s%, i%)
Dim j%, fnum%
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = App.path & "\..\data\"
  flNm = Mid(Date, 4, 2) & Mid(Date, 7, 2) & Trim(Str(i_s)) & "d.mpr"
  Open fDir & flNm For Output As #fnum
  Write #fnum, Date
  Write #fnum, Time
  Write #fnum, i
  For j = 0 To i
    Write #fnum, atemp(j, 0), atemp(j, 1), atemp(j, 2), apre(j), aposi(j)
  Next j
  Close #fnum
End Sub
Public Sub ResDtLoad(fDir$, flNm$)
Dim j%, fnum%, i%
  fnum = FreeFile
  Open fDir & flNm For Input As #fnum
  Input #fnum, gDate
  Input #fnum, gTime
  Input #fnum, gGphDtNum
  i = gGphDtNum
  For j = 0 To i
    Input #fnum, atemp(j, 0), atemp(j, 1), atemp(j, 2), apre(j), aposi(j)
  Next j
  Close #fnum
End Sub
Public Sub ResFlLoad()
Dim fDir$, fname$, rflg%
    
    fname = gResFlName        '結果データファイル名
    fDir = gResFldir          'ディレクトリ
    rflg = False
    Call GenFile.SetCtrl("ファイル読込", "読込", "取消")
    Call GenFile.SetFile(cLoad, fDir, fname, "*.mpr")
    GenFile.Show vbModal
    Call GenFile.GetFile(rflg, fDir, fname)
    Set GenFile = Nothing
    If rflg Then
      Screen.MousePointer = 11
      '
      ResDtLoad fDir, fname
      gResFlName = fname      'コントロールデータファイル名
      gResFldir = fDir        'ディレクトリ
      '
      Screen.MousePointer = 0
    End If
End Sub
' ---------------------------------------------------------
Public Sub coxDtSet()       '　cox データの　保存用「1ラインの文字データ」変換
Dim i%, fnum%, l%
Dim dmy$, dt$, com$
Dim iaf%, ja%
Dim isub As Long
Dim jsub As Long
Dim ksub As Long
Dim lsub As Long

    For l = 0 To 2: gCoxDlDt(l) = hcomm(l): Next l
    l = 4: gCoxDlDt(l) = ptime    '測定時間
    l = 6: gCoxDlDt(l) = ytemp    '予備加熱温度
    l = 7
    '軸駆動制御コマンドの読込
    For i = 0 To 100
      l = l + 1
      dt = gCoxDlDt(l)
      Mid(dt, 1, 2) = Right("  " & Str(seg_num(i)), 2)
      Mid(dt, 4, 4) = Right("    " & Str(ic(i)), 4)
      Mid(dt, 9, 9) = Right("         " & Format(z(i), "0.000"), 9)
      Mid(dt, 19, 10) = Right("        " & Format(vel(i), "0.00"), 10)
      Mid(dt, 30, 8) = Right("      " & Str(pres(i)), 8)
      Mid(dt, 39, 8) = Right("      " & Format(t0(i), "0.0"), 8)
      Mid(dt, 48, 6) = Right("      " & Format(p(i), "0.0"), 6)
      '
      gCoxDlDt(l) = dt
      If ic(i) = 9 Then Exit For
    Next i
    'データを読み取る
    l = l + 1
    '
    ja = 0
    For i = 0 To 200
      isub = sisub(i)
      jsub = sjsub(i)
      ksub = sksub(i)
      lsub = slsub(i)
      l = l + 1
      dt = gCoxDlDt(l)
      scom(i) = Mid(dt, 1, 2)
      Mid(dt, 4, 5) = Right("     " & Format(isub, "0"), 5)
      com = Left(scom(i), 1)
      Select Case com
      Case "S", "L"                    ' 2006.5.15 "L" 追加 s.f
        Mid(dt, 10, 5) = Right("     " & Format(jsub, "0"), 5)
        Mid(dt, 16, 5) = Right("     " & Format(ksub, "0"), 5)
        Mid(dt, 22, 5) = Right("     " & Format(lsub, "0"), 5)
      Case "J"

      Case "H"
      
      Case "P"

      Case "E"
        Exit For
      End Select

      gCoxDlDt(l) = dt
    Next i
'  -- 温度係数、肉厚補正データ
    l = l + 1   ' コメント行
    l = l + 1
    gCoxDlDt(l) = "  " & Format(T_keisuCont(0), "0.000") & ",  " & Format(T_keisuCont(1), "0.000")
    l = l + 1
    dt = "  " & Format(T_keisu(0), "0.000")
    For i = 1 To 4
      dt = dt & ",  " & Format(T_keisu(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
'
    l = l + 1
    dt = "  " & Format(T_keisu(5), "0.000")
    For i = 6 To 9
      dt = dt & ",  " & Format(T_keisu(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
 '
    l = l + 1  '  コメント行
    l = l + 1
    gCoxDlDt(l) = "  " & Format(Z3_HoseiCont(0), "0.000") & ",  " & Format(Z3_HoseiCont(1), "0.000") & ",  " & Format(Z3_HoseiCont(2), "0.000")
    l = l + 1
    dt = "  " & Format(Z3_Hosei(0), "0.000")
    For i = 1 To 4
      dt = dt & ",  " & Format(Z3_Hosei(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
'
    l = l + 1
    dt = "  " & Format(Z3_Hosei(5), "0.000")
    For i = 6 To 9
      dt = dt & ",  " & Format(Z3_Hosei(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
'
'
  '  加圧時間制御ｍａｘ、　ｍｉｎの書き込み
    l = l + 1   ' コメント行
    l = l + 1
    dt = "  " & Format(DkatJ(1), "000.0") & ",  " & Format(DkatJ(0), "000.0")
    gCoxDlDt(l) = dt
'
  '  型No.　データ　の書き込み
    l = l + 1   ' コメント行
    l = l + 1
    dt = "  " & kataNo(0)
    For i = 1 To 8
        dt = dt + ",  " & kataNo(i)
    Next i
    gCoxDlDt(l) = dt
'
'  加圧時間　＆　Ｃｐ値　Ａｌａｒｍ　ｍａｘ、　ｍｉｎの書き込み
    l = l + 1   ' コメント行
    l = l + 1
    dt = "  " & Format(AkatJ(1), "0") & ",  " & Format(AkatJ(0), "0")
    gCoxDlDt(l) = dt
'
    l = l + 1
    dt = "  " & Format(Acp(1), "0.000") & ",  " & Format(Acp(0), "0.000")
    gCoxDlDt(l) = dt
''
  Close fnum
End Sub
Public Sub coxDtSave(fl$)
Dim l%, fnum%
  fnum = FreeFile
  Open fl For Output As #fnum
    For l = 0 To gCoxFlDtMax
      Print #fnum, gCoxDlDt(l)
    Next l
  Close #fnum
End Sub

Private Sub DebugData()
Dim i%
Dim z!, p!, t!, x!
'
  For i = 0 To ResDtSize
    TPass(i) = i                '経過時間(秒)
    ZAxis(i) = 50 + 40 * Sin(i / 57.325)              '座標（Z-軸）
    Press(i) = i / 2000              '型締圧
    Templ(i) = 500 + 100 * Sin(i / 57.325)       '型温度
  Next i
End Sub

Public Sub BoardInit()
Dim flg%
    flg = 1
    Select Case flg
    Case 0
        BrdFlg = "OFF"
    Case 1
        BrdFlg = "ON"
        '--------------- D/A Board
        DeviceDaName
        'DvcDaOpen
        '--------------- A/D Board
        DvcAdOpen
        DeviceAdName
        '--------------- DIO Board
        DvcDioOpen
        '--------------- C-870V1
        Ready_Wait
        C870Open
    End Select
End Sub
Public Sub BoardClose()
Dim flg%
    flg = 1
    Select Case flg
    Case 0
        BrdFlg = "OFF"
    Case 1
        BrdFlg = "ON"
        '--------------- D/A Board
        'DeviceDaName
        DvcDaClose
        '--------------- A/D Board
        DvcAdClose
        'DeviceAdName
        '--------------- DIO Board
        'DvcDioClose
        '--------------- C-870V1
        Ready_Wait
        C870Close
    End Select
End Sub

Public Sub rozFileLoad()
Dim i%, fnum%
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = InitStr(3)
  flNm = InitStr(2)
  Open fDir & flNm For Input As #fnum
    Input #fnum, pv_ch                  '位置・速度モード切換点
    Input #fnum, roz(0), roz(1)         '突当成形ﾊﾟﾗﾒｰﾀ　幅、時間 (時間max180）
    Input #fnum, VccLw, VccHi           'ピラニゲージ用
    Input #fnum, gM_sa, gM_p, gM_lim    '速度制御のパラメータ
    Input #fnum, gHiSpeed, gLwSpeed     '手動の位置制御速度
    Input #fnum, r_pres_kousei          '圧力読取値　０校正
  Close #fnum
'gM_sa!     'メニューの速度制御の時の/* 設定圧力 */
'gM_p!      'メニューの速度制御の時の/* 設定Ｐ値 */
'gM_lim!    'メニューの速度制御の時の/* 設定リミット値 */
End Sub
Public Sub rozFileSave()
Dim i%, fnum%
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = InitStr(3)
  flNm = InitStr(2)
  Open fDir & flNm For Output As #fnum
    Write #fnum, pv_ch
    Write #fnum, roz(0), roz(1)        '突当成形ﾊﾟﾗﾒｰﾀ　幅、時間
    Write #fnum, VccLw, VccHi
    Write #fnum, gM_sa, gM_p, gM_lim
    Write #fnum, gHiSpeed, gLwSpeed    '手動の位置制御速度
    Write #fnum, r_pres_kousei          '圧力読取値　０校正
  Close #fnum
End Sub
Public Sub ExecMemo(DDir$, flNm$)
Dim ExecFl$, fl$
Dim r!
  fl = DDir$ & flNm
  ExecFl = "C:\WINDOWS\NOTEPAD.EXE " & fl
'-------- メモ帳でflを開く
  r = Shell(ExecFl, 1)
  AppActivate r, True     'メモ帳が閉じるまで待つ
End Sub
Public Function diffTime!(wTm1!, wTm0!)  '  '02.6.29  abs 外す   !入れる 10/4 sf
'Dim wTm0!, wTm1!
'-------------- ｛　wTm1（現在）－　wTm0(過去) ｝時間をSecで計算
  If wTm0 > wTm1 Then
    diffTime = wTm1 + H24Hr - wTm0
  Else
    diffTime = wTm1 - wTm0
    'diffTime = Abs(wTm1 - wTm0)
  End If
End Function

Public Function BitBSet%(dl%, bit%)
'
  BitBSet = dl Or (2 ^ bit%)

End Function
Public Function BitBReSet%(dl%, bit%)
'
  BitBReSet = dl And (&HFFFF - 2 ^ bit)

End Function
Public Function BitBTest%(dl%, bit%)
Dim sts%
'
  sts = 0
  If dl And 2 ^ bit Then sts = 1  '&h1
  BitBTest = sts
End Function
Public Sub cfileLoad()
Dim i%, fnum%
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = App.path & "\..\cont\"
  flNm = "cfile.con"
  Open fDir & flNm For Input As #fnum
    Input #fnum, gcoxFlName       'コントロールデータファイル名
    Input #fnum, gcoxFldir        'ディレクトリ
  Close #fnum
End Sub
Public Sub cfileSave()
Dim i%, fnum%
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = App.path & "\..\cont\"
  flNm = "cfile.con"
  Open fDir & flNm For Output As #fnum
    Write #fnum, gcoxFlName       'コントロールデータファイル名
    Write #fnum, gcoxFldir        'ディレクトリ
  Close #fnum
End Sub
Public Sub WaitSec(t As Single)
'単位 秒
Dim tm!, InTm!, NTm!
  tm = 0
  InTm = Timer
  Do
    NTm = Timer
    DoEvents
    If NTm >= InTm Then
      tm = NTm - InTm
    Else
      tm = H24Hr - InTm + NTm
    End If
    'If gDurPauseFlg <> 0 Then Exit Do
    If tm > t Then Exit Do
  Loop
End Sub

Public Sub SetErrMsg()
Dim ErrNo%, EmgArm%
  EmgArm = 0          '非常停止
  ErrNo = 0: gErrMsg$(EmgArm, ErrNo) = "System not ready" '
  ErrNo = 1: gErrMsg$(EmgArm, ErrNo) = "ＰＣ→非常停止" 'エラーメセージ
  ErrNo = 2: gErrMsg$(EmgArm, ErrNo) = "本体　非常停止ＳＷ"
  ErrNo = 3: gErrMsg$(EmgArm, ErrNo) = "制御盤　非常停止ＳＷ" '　’08.3　表示内容入替
  ErrNo = 4: gErrMsg$(EmgArm, ErrNo) = "高周波ＩＨ電源ＡＬＭ　成形室"
  ErrNo = 5: gErrMsg$(EmgArm, ErrNo) = "成形室上型ヒーターALM"
  ErrNo = 6: gErrMsg$(EmgArm, ErrNo) = "サーボモータ異常"
  ErrNo = 7: gErrMsg$(EmgArm, ErrNo) = "チャンバ圧異常"
  ErrNo = 8: gErrMsg$(EmgArm, ErrNo) = "ペルジャ圧異常"
  ErrNo = 9: gErrMsg$(EmgArm, ErrNo) = "成形室　温調器ＡＬＭ"
  ErrNo = 10: gErrMsg$(EmgArm, ErrNo) = "搬送中にORGが切れた"
  ErrNo = 11: gErrMsg$(EmgArm, ErrNo) = "予備加熱①ＨＦ電源ＡＬＭ"
  ErrNo = 12: gErrMsg$(EmgArm, ErrNo) = "予備加熱①温調器ＡＬＭ"
  ErrNo = 13: gErrMsg$(EmgArm, ErrNo) = "予備加熱②ＩＨ電源ＡＬＭ"
  ErrNo = 14: gErrMsg$(EmgArm, ErrNo) = "予備加熱②温調器ＡＬＭ"
  ErrNo = 15: gErrMsg$(EmgArm, ErrNo) = "成形室下型ヒーターALM"
  EmgArm = 1          'アラーム
  ErrNo = 0: gErrMsg$(EmgArm, ErrNo) = "ＡＬＭ　０　（未使用）" 'エラーメセージ
  ErrNo = 1: gErrMsg$(EmgArm, ErrNo) = "ペルジャ未到達" 'エラーメセージ
  ErrNo = 2: gErrMsg$(EmgArm, ErrNo) = "テーブル未到達"
  ErrNo = 3: gErrMsg$(EmgArm, ErrNo) = "パレット３未到達"
  ErrNo = 4: gErrMsg$(EmgArm, ErrNo) = "パレット４未到達"
  ErrNo = 5: gErrMsg$(EmgArm, ErrNo) = "パレット２未到達"
  ErrNo = 6: gErrMsg$(EmgArm, ErrNo) = "パレット１未到達"
  ErrNo = 7: gErrMsg$(EmgArm, ErrNo) = "ＡＬＭ７（未使用）"
  ErrNo = 8: gErrMsg$(EmgArm, ErrNo) = "成形室温度異常"
  ErrNo = 9: gErrMsg$(EmgArm, ErrNo) = "予備加熱２温度異常"
  ErrNo = 10: gErrMsg$(EmgArm, ErrNo) = "予備加熱１温度異常"
  ErrNo = 11: gErrMsg$(EmgArm, ErrNo) = "水冷却シリンダー下未到達"
  ErrNo = 12: gErrMsg$(EmgArm, ErrNo) = "真空未到達"
  ErrNo = 13: gErrMsg$(EmgArm, ErrNo) = "水冷却シリンダー上未到達"
  ErrNo = 14: gErrMsg$(EmgArm, ErrNo) = "予備加熱下移動未到達"
  ErrNo = 15: gErrMsg$(EmgArm, ErrNo) = "予備加熱上移動未到達"
End Sub
Public Sub DispCenter(frmObj As Form)
  Dim dmy As Long

  If frmObj.WindowState <> 0 Then frmObj.WindowState = 0
  dmy = Screen.Width - frmObj.Width
  If 1 < dmy Then
    frmObj.Left = dmy \ 2
  Else
    frmObj.Left = 0
  End If
  dmy = Screen.Height - frmObj.Height
  If 1 < dmy Then
    frmObj.Top = dmy \ 2
  Else
    frmObj.Top = 0
  End If
End Sub
