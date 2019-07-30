Attribute VB_Name = "PGM_S_Main"
Global VccDt!(0 To 10)        '真空計のデータ
Global VccCnt%                '真空計のデータ配列のカウンタ
Global VccTm0!
Global VccTm1!

Public Sub hijyou()
'/* 非常停止をシーケンサにかける */
'outp(DIO_P,0x10);   /* 非常停止をシーケンサへ依頼 */
  EmgFlg = True
  DioOut 5, 1     '非常停止
  SeikeiOFF       '成形OFF　待機中
  ServoOFF
  OrgOFF
'  VacuumOFF      '2006.12.21 削除 s.f
  HeatOFF
  CoolOFF
  FrmEmg.Show 1
End Sub

'Public Sub cal_pid(m_sa!, m_p!, m_lim!)
''  float  m_sa,     /* 設定圧力 */
''         m_p,      /* 設定Ｐ値 */
''         m_lim;    /* 設定リミット値 */
'Dim i%, nout%
'Dim pa!, per!
'  pa = r_pres()     '/* 圧力 */
'
'  If pa > 1000 Then '/* 指定圧力＋２００Ｋｇ以上で非常停止 */
'    hijyou
'    Exit Sub
'  End If
'
''/* ＰＩＤ演算 */
'
'  per = 5 * (m_sa - pa) * Abs(m_sa - pa) / (m_p * m_p)
'  If per > m_lim Then per = m_lim
'  If per < (-1 * m_lim) Then per = -1 * m_lim
'  inout = Int(40.95 * per) + &H800
'
'  'outp(ADPORT,(nout%256));
'  'outp(ADPORT+1,0x20|(nout/256));
'
'End Sub

Public Sub err_sign(ic%, id%)
Dim i%, ih%
'
' ic
' B0-B3 : 吸着エラー ～ サーボエラー
' B4    : 吸着エラー上下切換
' id
' エラー １～１２
  Err_ic = ic: Err_id = id
  'frmerr_sign.Show
  'Call frmerr_sign.DispErr
  
End Sub


Public Sub ErrBitRd(er1%, er2%, er3%)
Dim ch%
Dim er%(0 To 32), hdt%
'
  er1 = 0: er2 = 0: er3 = 0
  For ch = 5 To 8
    DioInput ch, hdt
    If hdt = 1 Then er1 = BitBSet(er1, ch - 5)
  Next ch
  For ch = 17 To 24
    DioInput ch, hdt
    If hdt = 1 Then er2 = BitBSet(er2, ch - 17)
  Next ch
  For ch = 25 To 32
    DioInput ch, hdt
    If hdt = 1 Then er3 = BitBSet(er3, ch - 25)
  Next ch
  '
End Sub
Public Function BitRd%(ch%)
Dim hdt%, sts%
'
  sts = 0
  Select Case ch
  Case 0
    For ch = 1 To 8
      DioInput ch, hdt
      If hdt = 1 Then sts = BitBSet(sts, ch - 1)
    Next ch
  Case 1
    For ch = 9 To 16
      DioInput ch, hdt
      If hdt = 1 Then sts = BitBSet(sts, ch - 9)
    Next ch
  Case 2
    For ch = 17 To 24
      DioInput ch, hdt
      If hdt = 1 Then sts = BitBSet(sts, ch - 17)
    Next ch
  Case 3
    For ch = 25 To 32
      DioInput ch, hdt
      If hdt = 1 Then sts = BitBSet(sts, ch - 25)
    Next ch
  End Select
  '
  BitRd = sts
  '
End Function

Public Sub outp(ch%, hdt%)
Dim sts%
  Select Case ch
  Case 0
    For ch = 1 To 8
      sts = BitBTest(hdt, ch)
      DioOut ch, sts
    Next ch
  Case 1
    For ch = 9 To 16
      sts = BitBTest(hdt, ch - 9)
      DioOut ch, sts
    Next ch
  Case 2
    For ch = 17 To 24
      sts = BitBTest(hdt, ch - 17)
      DioOut ch, sts
    Next ch
  Case 3
    For ch = 25 To 32
      sts = BitBTest(hdt, ch - 25)
      DioOut ch, sts
    Next ch
  End Select

End Sub
'Public Sub LS21S_Monitor()
'Dim dt!(0 To 4), i%
'Dim flg As Long
'  VccTm0 = Timer
'  If Int(VccTm0 * 10) = Int(VccTm1 * 10) Then Exit Sub
'  VccTm1 = VccTm0
'  flg = 0
'  'AdRead dt(), flg
'  '------ シーケンサで真空計の電源とＡＮＤ(2002.6.11)
'  'If dt(4) < VccHi Or gVumFlg = 1 Then
'  '  gVumFlg = 0                       '真空到達=1
'    VacuumON                           '真空到達信号発信
'  'Else
'    VacuumOFF                           '真空未到達信号発信
'  'End If
'  Exit Sub
'  '
'  If dt(4) <= 0.2 Then Exit Sub
'
'  VccCnt = VccCnt + 1
'  If VccCnt < 0 Then VccCnt = 0
'  If VccCnt > 3 Then VccCnt = 0
'  VccDt(VccCnt) = dt(4)
'  '---- 真空計　Ｌレンジ　１　～　３
'  '  真空ZERO　　　　　　　　　真空到達点
'  For i = 1 To 3
'    If VccLw < VccDt(i) And VccDt(i) < VccHi Then flg = flg + 1
'  Next i
'  If flg = 3 Or gVumFlg = 1 Then
'    gVumFlg = 0                       '真空到達=1
'    VacuumON                 '    '2006.12.21 削除 s.f
'    For i = 0 To 4: VccDt(i) = 10: Next i
'    WaitSec 1
'    VacuumOFF          '    '2006.12.21 削除 s.f
'  Else
'    VacuumOFF
'  End If
'End Sub


Public Function DispCtrlCommand$(i%)
Dim sdt$
    sdt = Right("     " & Format(i, "0"), 4)
    sdt = sdt & "  " & Right("     " & Format(seg_num(i), "0"), 4)   ' /* セグメント番号 */
    sdt = sdt & "  " & Right("     " & Format(ic(i), "0"), 4)        ' /* 制御方式 1,2,3,8,9 */
    sdt = sdt & "  " & Right("         " & Format(z(i), "0.000"), 7) ' /* 目標位置 */
    sdt = sdt & "  " & Right("         " & Format(vel(i), "0.0"), 7) ' /* 速度 */
    sdt = sdt & "  " & Right("       " & Format(pres(i), "0"), 6)    ' /* プレス圧力 */
    sdt = sdt & "  " & Right("     " & Format(t0(i), "0.0"), 4)      ' /* タイムアウト値 */
    sdt = sdt & "  " & Right("     " & Format(p(i), "0.0"), 4)       ' /* ＰＩＤ　Ｐ */
  DispCtrlCommand = sdt
End Function
