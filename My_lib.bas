Attribute VB_Name = "My_lib"

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Mylib
'   update: 2004.11.2  T係数関数　変更元へ戻す（方式１へ）　　s.f
'   update: 2004.9.26  T係数関数　変更　（方式２へ）　s.f
'   update: 2002.6.28  s.f. public sub cal_pid 追加
'   update: 2002.6.20 D/Aフルスケール変更(10V for 400kgf)
'   update: 2002.6.17 D/Aフルスケール変更02.6.17
'   update: 2005.11. 6 s.f   オーバーフロー対策　　long,doubleへ書き替え r_z!(),s_drive,setcm1
'   update: 2005.11.22 s.f   Melec C-870 counter動作バグ修正　コンペアカウンタ値セット時　符号反転　　setcm1
'   update: 2005.11.23 s.f   rstcm1 tsuika
'   update: 2005.11.26 s.f   定数の　＃化
'   update: 2005.12.23 s.f   longdata 計算　1行　→　3行
'   update: 2006. 5. 9 s.f    ppos = ppos & " r_z"
'   update: 2006. 5.14 s.f 　r_pres()の　DoEvents 　 forの外へ移動　s.f  ものすごく効く
'　　　　　　　　　　　　　　すべて抜くと　LS_TC　プログラム暴走する（LS_SCは　OK)’
'   update: 2006. 5.23 s.f 　cal_pid 変更
'   update: 2006. 7.12 s.f 　my_lib の　r_z!()　w1,w2,w3 long → integer
'   update: 2008. 4.14 s.f.  cal_pid　speed調整
'   update: 2008. 6. 2 s.f   Melec C-870 counter動作バグ修正　コンペアカウンタ値セット時　符号反転　　setcm1 azd=-ad * gDirect へ
'   update: 2008.11.17 s.f   cal_pid 「800kg以上で非常停止」　→　「１０００kg以上で非常停止」へ変更
'   update: 2009. 8.17 s.f   Timer2func 追加 timer overflow 対策
'　 update: 2012.04.15.s.f.　1ton越えの判断　１回→2回へ　　ＭｙＬｉｂ
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Timer2func!()          'Timer関数overflow対策
   Tm2f1 = Timer                       'Timerを２回読み　小さい方をとる（異常に大きい値をさける）
   Tm2f1 = Timer                       'Timerを２回読み　小さい方をとる（異常に大きい値をさける）
   Timer2func = Tm2f1                  '実際には、２回読み　２回目が小さければ２回目の値を取る。
   Tm2f2 = 0                           '
   Tm2f2 = Timer
   Tm2f2 = Timer
   If Tm2f2 < Tm2f1 Then Timer2func = Tm2f2
End Function
Public Function r_z!()
   Dim LongData As Long
'   Dim Longr_z As Long
'   Dim w1, w2, w3  As Long
   Dim w1, w2, w3  As Integer    ' 2006.7.12 s.f.OverFlow　対策
   If BrdFlg <> "ON" Then Exit Function
  '-------------------------- Z位置読み取り
   Ack = MPL_IRDrive(hDev, MplData, MplResult)   '現在位置ＡＤＤＲＥＳＳの表示
   w1 = MplData.MPL_Data(1)
   w2 = MplData.MPL_Data(2)
   w3 = MplData.MPL_Data(3)
      ppos = Left(ppos, 17) & " (r_z)"
   LongData = (w1 * idc65536)           '2005.11. 6 s.f  2005.12.23
   LongData = LongData + (w2 * idc256)  '2005.11. 6 s.f  2005.12.23
   LongData = LongData + w3             '2005.11. 6 s.f  2005.12.23
   If LongData > idc8388607 Then LongData = LongData - idc16777216
'   r_z = -LongData / gRev2Disp　　’　　LS
   r_z = LongData / gRev2Disp   '  '08.3.25  NQD
   '
   'If r_z > 0.1 Then OrgOFF      '原点LED　off  2002.10.9 KYOCERA
   '
End Function

Public Function r_pres!()
Dim i%, l%, ll%
Dim sumdt1!, sumd10!, dtrp!
'Dim dt!(0 To 4)
Dim dt!(0 To 7)
Dim adFlg As Long
  ppos = Left(ppos, 22) & " (r_pre)"
  sumdt10 = 0
'  DoEvents              '　2006.5.14　移動  2006.5.18 削除
  For l = 1 To 10
  sumdt1 = 0
   For ll = 1 To 10
    AdRead dt(), adFlg
    dtrp = dt(2) * 500#     'D/Aフルスケール変更 '08.4.22 NQD
    sumdt1 = sumdt1 + dtrp
   Next ll
   sumdt1 = sumdt1 / 10#
    If sumdt1 > 1000# Then
'        AdRead dt(), adFlg
'        dtrp = dt(2) * 500#     '1ton越え　２回確認へ　2012.4.15
'        If dtrp > 1000# Then
'            gemgmsg = gemgmsg + "r_pres" + Format(dtrp, "0.0")   '2012.0830
            gemgmsg = gemgmsg + "r_pres" + Format(sumdt1, "0.0")   '2012.0830 10回平均へ
            iFlg_hijyou = 6       '  非常停止 r_pres 1ton越え
'        End If
    End If
   sumdt10 = sumdt10 + sumdt1
'
'    DoEvents                 ' 2006.5.14 forの外へ移動　s.f  ものすごく効く
    'sumdt = sumdt + (dt(2) - 2.07667) * 223.8   '荷重 66kgで校正
'    sumdt = sumdt + dt(2) * 50#   'D/Aフルスケール変更02.7.31(10V for 500kgf)
    'sumdt = sumdt + dt(2) * 40  'D/Aフルスケール変更02.6.20(10V for 400kgf)
  Next l
  r_pres = sumdt10 / 10# - r_pres_kousei   '平均   ０次の項は、各機械で数値をみて決める（robo内のデータ）
End Function

Public Sub s_drive(az!, v!)
Dim k_puls As Long, hspd As Long
Dim sb As Double
Dim i%, sts%
Dim idt1 As Long, idt2 As Long, idt3 As Long
Dim ihd As Long
Dim sn  As Long
Dim pos, azd As Double

'2002.10.9 KYOCERA
  sts = PCTrnsChk
  If sts = 1 Then
    MsgBox "ＳＱ搬送中！　運転続行不可能", vbCritical + vbOKOnly, "致命的異常"
    End
  End If

'--------------- 速度の設定
  hspd = v * gRev2Disp / 60
'  If hspd > 400000 Then hspd = 400000  '02.5.11.sf
  If hspd > 800000 Then hspd = 800000
  If hspd < 77 Then hspd = 77
  
  Call MplDataSet(hspd, MplData)      'ＩＮＣＲＥＭＥＮＴＡＬ ＩＮＤＥＸ ＤＲＩＶＥ ＣＯＭＭＡＮＤ
  Ack = MPL_IWDrive(hDev, &H8, MplData, MplResult)      '　&H8　highspeed　設定
  
'--------------- パルス数の算出
  azd = az
  pos = r_z()
  k_puls = (azd - pos) * gRev2Disp + ddc05
  'If k_puls > 0 Then sn = 1 Else sn = -1
  'idt1 = Int(k_puls * sn / idc65536)
  'idt2 = Int((k_puls * sn - idt1 * idc65536) / idc256)
  'idt3 = k_puls * sn - idt1 * idc65536 - idt2 * idc256
'--------------- インクリメント動作
  Ready_Wait    'while((inp(AX_STS)&1)!=0);
  'Data = idt1: Ack = MPL_BWDriveData1(hDev, Data, MplResult)   '
  'Data = idt2: Ack = MPL_BWDriveData2(hDev, Data, MplResult)   '
  'Data = idt3: Ack = MPL_BWDriveData3(hDev, Data, MplResult)   '
  'cmd = &H14: Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)   '
  Call MplDataSet(k_puls, MplData)                    'ＩＮＣＲＥＭＥＮＴＡＬ ＩＮＤＥＸ ＤＲＩＶＥ ＣＯＭＭＡＮＤ
  Ack = MPL_IWDrive(hDev, &H14, MplData, MplResult)  ' &H14:　index　drive
End Sub
Public Sub rstcm1()
Dim zclear!
  zclear = -200#
  setcm1 zclear
End Sub
Public Sub setcm1(az!)
Dim k_puls As Long
Dim idt1, idt2, idt3, sn As Long
Dim i%
Dim azd As Double
'--------------- 到達パルス演算
  sn = 1
'  azd = -az          ' 何故だか解らないが　「－」で正常動作　　2005.11.22　ｓ.ｆ.
'  azd = -az * gDirect        ' 2008.5.  NQD立ち上げで　再度ｃｏｍｐcounter機能しない。　符号の問題！！　ｓ.ｆ.
  azd = az            ' 2008.9.15 NQD-2 立ち上げで　再度ｃｏｍｐcounter機能しない。　パルス数の符号は　正　が正しい！
  k_puls = azd * gRev2Disp + ddc05
'  idt1 = Int(k_puls * sn / idc65536)　　　　　　　　　　　　’　2005.11.22　　MPL_IWCounter　コマンドへ書替え
'  idt2 = Int((k_puls * sn - idt1 * idc65536) / idc256)
'  idt3 = k_puls * sn - idt1 * idc65536 - idt2 * idc256
'--------------- コンパレータ　１設定
  Ready_Wait    'while((inp(AX_STS)&1)!=0);
'  Data = idt1: Ack = MPL_BWCounterData1(hDev, Data, MplResult)   '
'  Data = idt2: Ack = MPL_BWCounterData2(hDev, Data, MplResult)   '
'  Data = idt3: Ack = MPL_BWCounterData3(hDev, Data, MplResult)   '
'  Cmd = &H1: Ack = MPL_BWCounterCommand(hDev, Cmd, MplResult)
   Call MplDataSet(k_puls, MplData)                    'ＩＮＣＲＥＭＥＮＴＡＬ ＩＮＤＥＸ ＤＲＩＶＥ ＣＯＭＭＡＮＤ
   Ack = MPL_IWCounter(hDev, &H1, MplData, MplResult)
End Sub
Public Sub Counter0()
Dim k_puls As Long
Dim i%, idt1!, idt2!, idt3!, sn%
'--------------- カウンタ０
  Ready_Wait    'while((inp(AX_STS)&1)!=0);
  Data = 0: Ack = MPL_BWCounterData1(hDev, Data, MplResult)   '
  Data = 0: Ack = MPL_BWCounterData2(hDev, Data, MplResult)   '
  Data = 0: Ack = MPL_BWCounterData3(hDev, Data, MplResult)   '
  Cmd = 0: Ack = MPL_BWCounterCommand(hDev, Cmd, MplResult)
End Sub
Public Sub cal_pid(m_sa!, m_p!, m_lim!)
'  float  m_sa,     /* 設定圧力 */
'         m_p,      /* 設定Ｐ値 */
'         m_lim;    /* 設定リミット値 */
Dim i%, ch%, m_p1!, m_lim1!
'Dim i%, nout%, ch%, v!    nout,v はGlobal宣言へ 2004.3.12
Dim pa!, per!       '/* float（単精度浮動小数点型)*/
  ppos = ppos + "csub"
'
' ' 無負荷時のスピードは　m_limで決まる！
'
  pa = r_pres()     '/* 圧力 */
'
  If ((pa > 1000#) Or (pa < -200#)) Then  '/* 1000Ｋｇ以上で非常停止 */ '081117
'  If ((pa > 800#) Or (pa < -100#)) Then  '/* 800Ｋｇ以上で非常停止 */ '080510
'  If pa > m_sa + 200# Then '/* 指定圧力 + 200Ｋｇ以上で非常停止 */
  gemgmsg = gemgmsg + "cal_pid" + Format(pa, "0.0")   '2012.0805
  hijyou                  ' 2012.0806.  -200以下　追加
    Exit Sub
  End If

'/* ＰＩＤ演算 */               ’　perの値の大きさは　加圧制御の細かさ
  ppos = ppos + "1"
'  per = 2 * 5 * (m_sa - pa) * Abs(m_sa - pa) / (m_p1 * m_p1)  ' 2008.4.14 NQD　speed調整
  per = 5 * (m_sa - pa) * Abs(m_sa - pa) / (m_p * m_p) ' 2008.4.14 NQD　speed調整
  If per > m_lim Then per = m_lim
  If per < (-1# * m_lim) Then per = -1# * m_lim     ' 2006.5.23 #追加
'
'  per = per * gDirect     'S.Mの回転方向 (+1 or -1)      ' 2008.3 NQD1-tsubaki
  per = -per               'S.Mの回転方向 (+1 or -1)      ' 2008.9.12 NQD2 -touei
'
  'nout = Int(40.95 * per) + &H800
  ppos = ppos + "2"
  nout = &H800 - Int(4.095 * per / 4#)      '  080912  touei
'  nout = &H800 - Int(40.95 * per)          '  0803    tsubaki
  ch = 1
'  v = 10# * (Int(4.095 * per / 4#) / 2048#)   ' 2005.11.26
  ppos = ppos + "3"
  DaOut ch, Hex(nout)
  
End Sub
Public Function T_keisu_cset!(t0cs!, tccs!)       ' 05.11.26　s.f.　　overflow 対策 「！」つける
' /*  新設定温度＝温度係数＊設定温度　　の　計算
' /* t00=　設定温度
' /* tc=　温度係数
'  Dim t0cs!, tccs!, abs0!
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  計算方式　１　　絶対零度からの　比例
  Dim abs0!
   abs0 = -273#
'
   T_keisu_cset = (t0cs - abs0) * tccs + abs0
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  計算方式　２　　温度係数値分だけ　シフト
'
'  Dim kijyun!, sa!
'
'  kijyun = 1#
'
'   T_keisu_cset = t0cs + (tccs - kijyun) * 100
'
End Function
Public Function T_keisu_cread!(t0cr!, tccr!)    ' 05.11.26　s.f.　　overflow 対策 「！」つける
' /*  新現在温度＝現在温度/温度係数　　の　計算
' /* t00=　設定温度
' /* tc=　温度係数
'  Dim t0cr!, tccr!, abs0!
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  計算方式　１　　絶対零度からの　比例
  Dim abs0!
'
   abs0 = -273#
'
   T_keisu_cread = (t0cr - abs0) / tccr + abs0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  計算方式　２　　温度係数値分だけ　シフト
'
'  Dim kijyun!, sa!
'  kijyun = 1#
'
'    T_keisu_cread = t0cr - (tccr - kijyun) * 100
'
End Function

