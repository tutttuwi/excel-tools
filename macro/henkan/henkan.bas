'----------------
'半角カタカナ変換
'----------------
Function hankana(str As String) As String
hankana = StrConv(str, 24)
End Function

'----------------
'全角ひらかな変換
'----------------
Function hira(str As String) As String
hira = Replace(StrConv(str, 36), "ヴ", "う゛")
End Function

'----------------
'全角カタカナ変換
'----------------
Function kana(str As String) As String
kana = StrConv(str, 20)
End Function

'--------------------
'カナをローマ字に変換
'--------------------
Function roma(ByVal kana As String, Optional ByVal capital As Boolean = False) As String
  Dim i As Integer, retStr As String
  Dim ts As Boolean, tmp1 As String, tmp2 As String
  Dim Cnv2(1 To 92, 1 To 2) As String, Cnv1(1 To 87, 1 To 2) As String
  Cnv2(1, 1) = "イェ": Cnv2(1, 2) = "ye"
  Cnv2(2, 1) = "ヴァ": Cnv2(2, 2) = "va"
  Cnv2(3, 1) = "ウィ": Cnv2(3, 2) = "wi"
  Cnv2(4, 1) = "ヴィ": Cnv2(4, 2) = "vi"
  Cnv2(5, 1) = "ウェ": Cnv2(5, 2) = "we"
  Cnv2(6, 1) = "ヴェ": Cnv2(6, 2) = "ve"
  Cnv2(7, 1) = "ヴォ": Cnv2(7, 2) = "vo"
  Cnv2(8, 1) = "キィ": Cnv2(8, 2) = "kyi"
  Cnv2(9, 1) = "ギィ": Cnv2(9, 2) = "gyi"
  Cnv2(10, 1) = "キェ": Cnv2(10, 2) = "kye"
  Cnv2(11, 1) = "ギェ": Cnv2(11, 2) = "gye"
  Cnv2(12, 1) = "キャ": Cnv2(12, 2) = "kya"
  Cnv2(13, 1) = "ギャ": Cnv2(13, 2) = "gya"
  Cnv2(14, 1) = "キュ": Cnv2(14, 2) = "kyu"
  Cnv2(15, 1) = "ギュ": Cnv2(15, 2) = "gyu"
  Cnv2(16, 1) = "キョ": Cnv2(16, 2) = "kyo"
  Cnv2(17, 1) = "ギョ": Cnv2(17, 2) = "gyo"
  Cnv2(18, 1) = "クァ": Cnv2(18, 2) = "kwa"
  Cnv2(19, 1) = "グァ": Cnv2(19, 2) = "gwa"
  Cnv2(20, 1) = "シィ": Cnv2(20, 2) = "syi"
  Cnv2(21, 1) = "ジィ": Cnv2(21, 2) = "zyi"
  Cnv2(22, 1) = "シェ": Cnv2(22, 2) = "she"
  Cnv2(23, 1) = "ジェ": Cnv2(23, 2) = "je"
  Cnv2(24, 1) = "シャ": Cnv2(24, 2) = "sha"
  Cnv2(25, 1) = "ジャ": Cnv2(25, 2) = "ja"
  Cnv2(26, 1) = "シュ": Cnv2(26, 2) = "shu"
  Cnv2(27, 1) = "ジュ": Cnv2(27, 2) = "ju"
  Cnv2(28, 1) = "ショ": Cnv2(28, 2) = "sho"
  Cnv2(29, 1) = "ジョ": Cnv2(29, 2) = "jo"
  Cnv2(30, 1) = "チィ": Cnv2(30, 2) = "tyi"
  Cnv2(31, 1) = "ヂィ": Cnv2(31, 2) = "dyi"
  Cnv2(32, 1) = "チェ": Cnv2(32, 2) = "che"
  Cnv2(33, 1) = "ヂェ": Cnv2(33, 2) = "dye"
  Cnv2(34, 1) = "チャ": Cnv2(34, 2) = "cha"
  Cnv2(35, 1) = "ヂャ": Cnv2(35, 2) = "dya"
  Cnv2(36, 1) = "チュ": Cnv2(36, 2) = "chu"
  Cnv2(37, 1) = "ヂュ": Cnv2(37, 2) = "dyu"
  Cnv2(38, 1) = "チョ": Cnv2(38, 2) = "cho"
  Cnv2(39, 1) = "ヂョ": Cnv2(39, 2) = "dyo"
  Cnv2(40, 1) = "ツァ": Cnv2(40, 2) = "tsa"
  Cnv2(41, 1) = "ツィ": Cnv2(41, 2) = "tsi"
  Cnv2(42, 1) = "ツェ": Cnv2(42, 2) = "tse"
  Cnv2(43, 1) = "ツォ": Cnv2(43, 2) = "tso"
  Cnv2(44, 1) = "ティ": Cnv2(44, 2) = "thi"
  Cnv2(45, 1) = "ディ": Cnv2(45, 2) = "dhi"
  Cnv2(46, 1) = "テェ": Cnv2(46, 2) = "the"
  Cnv2(47, 1) = "デェ": Cnv2(47, 2) = "dhe"
  Cnv2(48, 1) = "テャ": Cnv2(48, 2) = "tha"
  Cnv2(49, 1) = "デャ": Cnv2(49, 2) = "dha"
  Cnv2(50, 1) = "テュ": Cnv2(50, 2) = "thu"
  Cnv2(51, 1) = "デュ": Cnv2(51, 2) = "dhu"
  Cnv2(52, 1) = "テョ": Cnv2(52, 2) = "tho"
  Cnv2(53, 1) = "デョ": Cnv2(53, 2) = "dho"
  Cnv2(54, 1) = "トゥ": Cnv2(54, 2) = "twu"
  Cnv2(55, 1) = "ドゥ": Cnv2(55, 2) = "dwu"
  Cnv2(56, 1) = "ニィ": Cnv2(56, 2) = "nyi"
  Cnv2(57, 1) = "ニェ": Cnv2(57, 2) = "nye"
  Cnv2(58, 1) = "ニャ": Cnv2(58, 2) = "nya"
  Cnv2(59, 1) = "ニュ": Cnv2(59, 2) = "nyu"
  Cnv2(60, 1) = "ニョ": Cnv2(60, 2) = "nyo"
  Cnv2(61, 1) = "ヒィ": Cnv2(61, 2) = "hyi"
  Cnv2(62, 1) = "ビィ": Cnv2(62, 2) = "byi"
  Cnv2(63, 1) = "ピィ": Cnv2(63, 2) = "pyi"
  Cnv2(64, 1) = "ヒェ": Cnv2(64, 2) = "hye"
  Cnv2(65, 1) = "ビェ": Cnv2(65, 2) = "bye"
  Cnv2(66, 1) = "ピェ": Cnv2(66, 2) = "pye"
  Cnv2(67, 1) = "ヒャ": Cnv2(67, 2) = "hya"
  Cnv2(68, 1) = "ビャ": Cnv2(68, 2) = "bya"
  Cnv2(69, 1) = "ピャ": Cnv2(69, 2) = "pya"
  Cnv2(70, 1) = "ヒュ": Cnv2(70, 2) = "hyu"
  Cnv2(71, 1) = "ビュ": Cnv2(71, 2) = "byu"
  Cnv2(72, 1) = "ピュ": Cnv2(72, 2) = "pyu"
  Cnv2(73, 1) = "ヒョ": Cnv2(73, 2) = "hyo"
  Cnv2(74, 1) = "ビョ": Cnv2(74, 2) = "byo"
  Cnv2(75, 1) = "ピョ": Cnv2(75, 2) = "pyo"
  Cnv2(76, 1) = "ファ": Cnv2(76, 2) = "fa"
  Cnv2(77, 1) = "フィ": Cnv2(77, 2) = "fi"
  Cnv2(78, 1) = "フェ": Cnv2(78, 2) = "fe"
  Cnv2(79, 1) = "フォ": Cnv2(79, 2) = "fo"
  Cnv2(80, 1) = "フャ": Cnv2(80, 2) = "fya"
  Cnv2(81, 1) = "フュ": Cnv2(81, 2) = "fyu"
  Cnv2(82, 1) = "フョ": Cnv2(82, 2) = "fyo"
  Cnv2(83, 1) = "ミィ": Cnv2(83, 2) = "myi"
  Cnv2(84, 1) = "ミェ": Cnv2(84, 2) = "mye"
  Cnv2(85, 1) = "ミャ": Cnv2(85, 2) = "mya"
  Cnv2(86, 1) = "ミュ": Cnv2(86, 2) = "myu"
  Cnv2(87, 1) = "ミョ": Cnv2(87, 2) = "myo"
  Cnv2(88, 1) = "リィ": Cnv2(88, 2) = "ryi"
  Cnv2(89, 1) = "リェ": Cnv2(89, 2) = "rye"
  Cnv2(90, 1) = "リャ": Cnv2(90, 2) = "rya"
  Cnv2(91, 1) = "リュ": Cnv2(91, 2) = "ryu"
  Cnv2(92, 1) = "リョ": Cnv2(92, 2) = "ryo"
  Cnv1(1, 1) = "ァ": Cnv1(1, 2) = "la"
  Cnv1(2, 1) = "ア": Cnv1(2, 2) = "a"
  Cnv1(3, 1) = "ィ": Cnv1(3, 2) = "li"
  Cnv1(4, 1) = "イ": Cnv1(4, 2) = "i"
  Cnv1(5, 1) = "ゥ": Cnv1(5, 2) = "lu"
  Cnv1(6, 1) = "ウ": Cnv1(6, 2) = "u"
  Cnv1(7, 1) = "ヴ": Cnv1(7, 2) = "vu"
  Cnv1(8, 1) = "ェ": Cnv1(8, 2) = "le"
  Cnv1(9, 1) = "エ": Cnv1(9, 2) = "e"
  Cnv1(10, 1) = "ォ": Cnv1(10, 2) = "lo"
  Cnv1(11, 1) = "オ": Cnv1(11, 2) = "o"
  Cnv1(12, 1) = "ヵ": Cnv1(12, 2) = "lka"
  Cnv1(13, 1) = "カ": Cnv1(13, 2) = "ka"
  Cnv1(14, 1) = "ガ": Cnv1(14, 2) = "ga"
  Cnv1(15, 1) = "キ": Cnv1(15, 2) = "ki"
  Cnv1(16, 1) = "ギ": Cnv1(16, 2) = "gi"
  Cnv1(17, 1) = "ク": Cnv1(17, 2) = "ku"
  Cnv1(18, 1) = "グ": Cnv1(18, 2) = "gu"
  Cnv1(19, 1) = "ヶ": Cnv1(19, 2) = "lke"
  Cnv1(20, 1) = "ケ": Cnv1(20, 2) = "ke"
  Cnv1(21, 1) = "ゲ": Cnv1(21, 2) = "ge"
  Cnv1(22, 1) = "コ": Cnv1(22, 2) = "ko"
  Cnv1(23, 1) = "ゴ": Cnv1(23, 2) = "go"
  Cnv1(24, 1) = "サ": Cnv1(24, 2) = "sa"
  Cnv1(25, 1) = "ザ": Cnv1(25, 2) = "za"
  Cnv1(26, 1) = "シ": Cnv1(26, 2) = "shi"
  Cnv1(27, 1) = "ジ": Cnv1(27, 2) = "ji"
  Cnv1(28, 1) = "ス": Cnv1(28, 2) = "su"
  Cnv1(29, 1) = "ズ": Cnv1(29, 2) = "zu"
  Cnv1(30, 1) = "セ": Cnv1(30, 2) = "se"
  Cnv1(31, 1) = "ゼ": Cnv1(31, 2) = "ze"
  Cnv1(32, 1) = "ソ": Cnv1(32, 2) = "so"
  Cnv1(33, 1) = "ゾ": Cnv1(33, 2) = "zo"
  Cnv1(34, 1) = "タ": Cnv1(34, 2) = "ta"
  Cnv1(35, 1) = "ダ": Cnv1(35, 2) = "da"
  Cnv1(36, 1) = "チ": Cnv1(36, 2) = "chi"
  Cnv1(37, 1) = "ヂ": Cnv1(37, 2) = "di"
  Cnv1(38, 1) = "ツ": Cnv1(38, 2) = "tsu"
  Cnv1(39, 1) = "ヅ": Cnv1(39, 2) = "du"
  Cnv1(40, 1) = "テ": Cnv1(40, 2) = "te"
  Cnv1(41, 1) = "デ": Cnv1(41, 2) = "de"
  Cnv1(42, 1) = "ト": Cnv1(42, 2) = "to"
  Cnv1(43, 1) = "ド": Cnv1(43, 2) = "do"
  Cnv1(44, 1) = "ナ": Cnv1(44, 2) = "na"
  Cnv1(45, 1) = "ニ": Cnv1(45, 2) = "ni"
  Cnv1(46, 1) = "ヌ": Cnv1(46, 2) = "nu"
  Cnv1(47, 1) = "ネ": Cnv1(47, 2) = "ne"
  Cnv1(48, 1) = "ノ": Cnv1(48, 2) = "no"
  Cnv1(49, 1) = "ハ": Cnv1(49, 2) = "ha"
  Cnv1(50, 1) = "バ": Cnv1(50, 2) = "ba"
  Cnv1(51, 1) = "パ": Cnv1(51, 2) = "pa"
  Cnv1(52, 1) = "ヒ": Cnv1(52, 2) = "hi"
  Cnv1(53, 1) = "ビ": Cnv1(53, 2) = "bi"
  Cnv1(54, 1) = "ピ": Cnv1(54, 2) = "pi"
  Cnv1(55, 1) = "フ": Cnv1(55, 2) = "fu"
  Cnv1(56, 1) = "ブ": Cnv1(56, 2) = "bu"
  Cnv1(57, 1) = "プ": Cnv1(57, 2) = "pu"
  Cnv1(58, 1) = "ヘ": Cnv1(58, 2) = "he"
  Cnv1(59, 1) = "ベ": Cnv1(59, 2) = "be"
  Cnv1(60, 1) = "ペ": Cnv1(60, 2) = "pe"
  Cnv1(61, 1) = "ホ": Cnv1(61, 2) = "ho"
  Cnv1(62, 1) = "ボ": Cnv1(62, 2) = "bo"
  Cnv1(63, 1) = "ポ": Cnv1(63, 2) = "po"
  Cnv1(64, 1) = "マ": Cnv1(64, 2) = "ma"
  Cnv1(65, 1) = "ミ": Cnv1(65, 2) = "mi"
  Cnv1(66, 1) = "ム": Cnv1(66, 2) = "mu"
  Cnv1(67, 1) = "メ": Cnv1(67, 2) = "me"
  Cnv1(68, 1) = "モ": Cnv1(68, 2) = "mo"
  Cnv1(69, 1) = "ャ": Cnv1(69, 2) = "lya"
  Cnv1(70, 1) = "ヤ": Cnv1(70, 2) = "ya"
  Cnv1(71, 1) = "ュ": Cnv1(71, 2) = "lyu"
  Cnv1(72, 1) = "ユ": Cnv1(72, 2) = "yu"
  Cnv1(73, 1) = "ョ": Cnv1(73, 2) = "lyo"
  Cnv1(74, 1) = "ヨ": Cnv1(74, 2) = "yo"
  Cnv1(75, 1) = "ラ": Cnv1(75, 2) = "ra"
  Cnv1(76, 1) = "リ": Cnv1(76, 2) = "ri"
  Cnv1(77, 1) = "ル": Cnv1(77, 2) = "ru"
  Cnv1(78, 1) = "レ": Cnv1(78, 2) = "re"
  Cnv1(79, 1) = "ロ": Cnv1(79, 2) = "ro"
  Cnv1(80, 1) = "ヮ": Cnv1(80, 2) = "lwa"
  Cnv1(81, 1) = "ワ": Cnv1(81, 2) = "wa"
  Cnv1(82, 1) = "ワ": Cnv1(82, 2) = "wa"
  Cnv1(83, 1) = "ヰ": Cnv1(83, 2) = "wyi"
  Cnv1(84, 1) = "ヱ": Cnv1(84, 2) = "wye"
  Cnv1(85, 1) = "ヲ": Cnv1(85, 2) = "wo"
  Cnv1(86, 1) = "ン": Cnv1(86, 2) = "nn"
  Cnv1(87, 1) = "ー": Cnv1(87, 2) = "-"
  kana = StrConv(kana, vbKatakana Or vbWide)
  retStr = "": i = 1: ts = False
  Do While i <= Len(kana)
    tmp2 = "": tmp1 = ""
    On Error Resume Next
    tmp2 = Application.WorksheetFunction.VLookup(Mid(kana, i, 2), Cnv2(), 2, False)
    tmp1 = Application.WorksheetFunction.VLookup(Mid(kana, i, 1), Cnv1(), 2, False)
    On Error GoTo 0
    If tmp2 <> "" Then
      If ts Then
        retStr = retStr & IIf(capital, StrConv(Left(tmp2, 1), 1), Left(tmp2, 1))
      End If
      retStr = retStr & IIf(capital, StrConv(tmp2, 1), tmp2)
      i = i + 2
      ts = False
    ElseIf tmp1 <> "" Then
      If ts Then
        retStr = retStr & IIf(capital, StrConv(Left(tmp1, 1), 1), Left(tmp1, 1))
      End If
      retStr = retStr & IIf(capital, StrConv(tmp1, 1), tmp1)
      i = i + 1
      ts = False
    ElseIf Mid(kana, i, 1) = "ッ" Then
      If ts Then
         retStr = retStr & IIf(capital, "LTSU", "ltsu")
      End If
      ts = True
      i = i + 1
    Else
      retStr = retStr & Mid(kana, i, 1)
      i = i + 1
    End If
  Loop
  If ts Then
     retStr = retStr & IIf(capital, "LTSU", "ltsu")
  End If
  
  roma = retStr
End Function

'------------------------
'ローマ字を全角カナに変換
'------------------------
Function roma2kana(ByVal roma As String, Optional ByVal hiragana As Boolean = False) As String
  Dim i As Integer, j1 As Integer, j2 As Integer, k As Integer, index As Integer
  Dim conv2 As Variant, conv1 As Variant, conv0 As Variant
  Dim kanatbl(1 To 54), Pre As String
  Dim retStr As String
  conv2 = Array("by", "ch", "cy", "dh", _
          "dw", "dy", "fy", "gw", "gy", "hy", "jy", "kw", "ky", "lk", _
          "lt", "lw", "ly", "my", "nn", "ny", "py", "ry", "sh", "sy", _
          "th", "ts", "tw", "ty", "wy", "xk", "xt", "xw", "xy", "zy")
  conv1 = Array("b", "d", "f", "g", "h", "j", "k", "l", "m", "n", _
          "p", "r", "s", "t", "v", "w", "x", "y", "z")
  conv0 = Array("a", "i", "u", "e", "o")
  kanatbl(1) = Array("ア", "イ", "ウ", "エ", "オ")
  kanatbl(2) = Array("バ", "ビ", "ブ", "ベ", "ボ")
  kanatbl(3) = Array("ダ", "ヂ", "ヅ", "デ", "ド")
  kanatbl(4) = Array("ファ", "フィ", "フ", "フェ", "フォ")
  kanatbl(5) = Array("ガ", "ギ", "グ", "ゲ", "ゴ")
  kanatbl(6) = Array("ハ", "ヒ", "フ", "ヘ", "ホ")
  kanatbl(7) = Array("ジャ", "ジ", "ジュ", "ジェ", "ジョ")
  kanatbl(8) = Array("カ", "キ", "ク", "ケ", "コ")
  kanatbl(9) = Array("ァ", "ィ", "ゥ", "ェ", "ォ")
  kanatbl(10) = Array("マ", "ミ", "ム", "メ", "モ")
  kanatbl(11) = Array("ナ", "ニ", "ヌ", "ネ", "ノ")
  kanatbl(12) = Array("パ", "ピ", "プ", "ペ", "ポ")
  kanatbl(13) = Array("ラ", "リ", "ル", "レ", "ロ")
  kanatbl(14) = Array("サ", "シ", "ス", "セ", "ソ")
  kanatbl(15) = Array("タ", "チ", "ツ", "テ", "ト")
  kanatbl(16) = Array("ヴァ", "ヴィ", "ヴ", "ヴェ", "ヴォ")
  kanatbl(17) = Array("ワ", "ウィ", "ウ", "ウェ", "ヲ")
  kanatbl(18) = Array("ァ", "ィ", "ゥ", "ェ", "ォ")
  kanatbl(19) = Array("ヤ", "イ", "ユ", "イェ", "ヨ")
  kanatbl(20) = Array("ザ", "ジ", "ズ", "ゼ", "ゾ")
  kanatbl(21) = Array("ビャ", "ビィ", "ビュ", "ビェ", "ビョ")
  kanatbl(22) = Array("チャ", "チ", "チュ", "チェ", "チョ")
  kanatbl(23) = Array("チャ", "チィ", "チュ", "チェ", "チョ")
  kanatbl(24) = Array("デャ", "ディ", "デュ", "デェ", "デョ")
  kanatbl(25) = Array("dwa", "dwi", "ドゥ", "dwe", "dwo")
  kanatbl(26) = Array("ヂャ", "ヂィ", "ヂュ", "ヂェ", "ヂョ")
  kanatbl(27) = Array("フャ", "フィ", "フュ", "フェ", "フョ")
  kanatbl(28) = Array("グァ", "gwi", "gwu", "gwe", "gwo")
  kanatbl(29) = Array("ギャ", "ギィ", "ギュ", "ギェ", "ギョ")
  kanatbl(30) = Array("ヒャ", "ヒィ", "ヒュ", "ヒェ", "ヒョ")
  kanatbl(31) = Array("ジャ", "ジィ", "ジュ", "ジェ", "ジョ")
  kanatbl(32) = Array("クァ", "kwi", "kwu", "kwe", "kwo")
  kanatbl(33) = Array("キャ", "キィ", "キュ", "キェ", "キョ")
  kanatbl(34) = Array("ヵ", "lki", "lku", "ヶ", "lko")
  kanatbl(35) = Array("lta", "lti", "ッ", "lte", "lto")
  kanatbl(36) = Array("ヮ", "lwi", "lwu", "lwe", "lwo")
  kanatbl(37) = Array("ャ", "ィ", "ュ", "ェ", "ョ")
  kanatbl(38) = Array("ミャ", "ミィ", "ミュ", "ミェ", "ミョ")
  kanatbl(39) = Array("ンア", "ンイ", "ンウ", "ンエ", "ンオ")
  kanatbl(40) = Array("ニャ", "ニィ", "ニュ", "ニェ", "ニョ")
  kanatbl(41) = Array("ピャ", "ピィ", "ピュ", "ピェ", "ピョ")
  kanatbl(42) = Array("リャ", "リィ", "リュ", "リェ", "リョ")
  kanatbl(43) = Array("シャ", "シ", "シュ", "シェ", "ショ")
  kanatbl(44) = Array("シャ", "シィ", "シュ", "シェ", "ショ")
  kanatbl(45) = Array("テャ", "ティ", "テュ", "テェ", "テョ")
  kanatbl(46) = Array("ツァ", "ツィ", "ツ", "ツェ", "ツォ")
  kanatbl(47) = Array("twa", "twi", "トゥ", "twe", "two")
  kanatbl(48) = Array("チャ", "チィ", "チュ", "チェ", "チョ")
  kanatbl(49) = Array("wya", "ヰ", "wyu", "ヱ", "wyo")
  kanatbl(50) = Array("ヵ", "xki", "xku", "ヶ", "xko")
  kanatbl(51) = Array("xta", "xti", "ッ", "xte", "xto")
  kanatbl(52) = Array("ヮ", "xwi", "xwu", "xwe", "xwo")
  kanatbl(53) = Array("ャ", "ィ", "ュ", "ェ", "ョ")
  kanatbl(54) = Array("ジャ", "ジィ", "ジュ", "ジェ", "ジョ")
  roma = StrConv(roma, vbNarrow Or vbLowerCase)
  retStr = "": Pre = "": i = 1: index = 1
  Do While i <= Len(roma)
    k = 0: j1 = 0: j2 = 0
    If Mid(roma, i, 1) Like "[a-z-]" Then
      On Error Resume Next
      k = Application.WorksheetFunction.Match(Mid(roma, i, 1), conv0, 0)
      On Error GoTo 0
      If k > 0 Then
        retStr = retStr & IIf(index = 1, Pre, "") & _
           IIf(hiragana, Replace(StrConv(kanatbl(index)(k - 1), 36), "ヴ", "う゛"), _
           StrConv(kanatbl(index)(k - 1), 20))
        Pre = "": i = i + 1: index = 1
      ElseIf k = 0 Then
        On Error Resume Next
        j2 = Application.WorksheetFunction.Match(Mid(roma, i, 2), conv2, 0)
        j1 = Application.WorksheetFunction.Match(Mid(roma, i, 1), conv1, 0)
        On Error GoTo 0
        If j2 > 0 Then j1 = 0
        index = 1 - (j2 > 0) * 19 + j2 + j1
        Select Case Pre
          Case Mid(roma, i, 1)
            retStr = retStr & IIf(hiragana, "っ", "ッ")
          Case "nn"
            retStr = retStr & IIf(hiragana, "ん", "ン")
          Case "-"
            retStr = retStr & IIf(hiragana, "ー", "ー")
          Case "n"
            retStr = retStr & IIf(hiragana, "ん", "ン")
          Case Else
            retStr = retStr & Pre
        End Select
        Pre = Mid(roma, i, IIf(j2, 2, 1))
        i = i + 1 + IIf(j2, 1, 0)
        If Pre = "lt" Or Pre = "xt" Then
          If Mid(roma, i, 2) = "su" Then
            retStr = retStr + IIf(hiragana, "っ", "ッ")
            i = i + 2
            Pre = ""
          End If
        End If
      End If
    Else
      retStr = retStr + IIf(Pre = "nn" Or Pre = "n", IIf(hiragana, "ん", "ン"), Pre)
      If Mid(roma, i, 1) <> "'" Then
         retStr = retStr + Mid(roma, i, 1)
      End If
      Pre = ""
      index = 1
      i = i + 1
    End If
  Loop
  roma2kana = retStr & IIf(Pre = "nn" Or Pre = "n" Or Pre = "n'", IIf(hiragana, "ん", "ン"), Pre)
End Function



