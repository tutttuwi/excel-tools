'----------------
'���p�J�^�J�i�ϊ�
'----------------
Function hankana(str As String) As String
hankana = StrConv(str, 24)
End Function

'----------------
'�S�p�Ђ炩�ȕϊ�
'----------------
Function hira(str As String) As String
hira = Replace(StrConv(str, 36), "��", "���J")
End Function

'----------------
'�S�p�J�^�J�i�ϊ�
'----------------
Function kana(str As String) As String
kana = StrConv(str, 20)
End Function

'--------------------
'�J�i�����[�}���ɕϊ�
'--------------------
Function roma(ByVal kana As String, Optional ByVal capital As Boolean = False) As String
  Dim i As Integer, retStr As String
  Dim ts As Boolean, tmp1 As String, tmp2 As String
  Dim Cnv2(1 To 92, 1 To 2) As String, Cnv1(1 To 87, 1 To 2) As String
  Cnv2(1, 1) = "�C�F": Cnv2(1, 2) = "ye"
  Cnv2(2, 1) = "���@": Cnv2(2, 2) = "va"
  Cnv2(3, 1) = "�E�B": Cnv2(3, 2) = "wi"
  Cnv2(4, 1) = "���B": Cnv2(4, 2) = "vi"
  Cnv2(5, 1) = "�E�F": Cnv2(5, 2) = "we"
  Cnv2(6, 1) = "���F": Cnv2(6, 2) = "ve"
  Cnv2(7, 1) = "���H": Cnv2(7, 2) = "vo"
  Cnv2(8, 1) = "�L�B": Cnv2(8, 2) = "kyi"
  Cnv2(9, 1) = "�M�B": Cnv2(9, 2) = "gyi"
  Cnv2(10, 1) = "�L�F": Cnv2(10, 2) = "kye"
  Cnv2(11, 1) = "�M�F": Cnv2(11, 2) = "gye"
  Cnv2(12, 1) = "�L��": Cnv2(12, 2) = "kya"
  Cnv2(13, 1) = "�M��": Cnv2(13, 2) = "gya"
  Cnv2(14, 1) = "�L��": Cnv2(14, 2) = "kyu"
  Cnv2(15, 1) = "�M��": Cnv2(15, 2) = "gyu"
  Cnv2(16, 1) = "�L��": Cnv2(16, 2) = "kyo"
  Cnv2(17, 1) = "�M��": Cnv2(17, 2) = "gyo"
  Cnv2(18, 1) = "�N�@": Cnv2(18, 2) = "kwa"
  Cnv2(19, 1) = "�O�@": Cnv2(19, 2) = "gwa"
  Cnv2(20, 1) = "�V�B": Cnv2(20, 2) = "syi"
  Cnv2(21, 1) = "�W�B": Cnv2(21, 2) = "zyi"
  Cnv2(22, 1) = "�V�F": Cnv2(22, 2) = "she"
  Cnv2(23, 1) = "�W�F": Cnv2(23, 2) = "je"
  Cnv2(24, 1) = "�V��": Cnv2(24, 2) = "sha"
  Cnv2(25, 1) = "�W��": Cnv2(25, 2) = "ja"
  Cnv2(26, 1) = "�V��": Cnv2(26, 2) = "shu"
  Cnv2(27, 1) = "�W��": Cnv2(27, 2) = "ju"
  Cnv2(28, 1) = "�V��": Cnv2(28, 2) = "sho"
  Cnv2(29, 1) = "�W��": Cnv2(29, 2) = "jo"
  Cnv2(30, 1) = "�`�B": Cnv2(30, 2) = "tyi"
  Cnv2(31, 1) = "�a�B": Cnv2(31, 2) = "dyi"
  Cnv2(32, 1) = "�`�F": Cnv2(32, 2) = "che"
  Cnv2(33, 1) = "�a�F": Cnv2(33, 2) = "dye"
  Cnv2(34, 1) = "�`��": Cnv2(34, 2) = "cha"
  Cnv2(35, 1) = "�a��": Cnv2(35, 2) = "dya"
  Cnv2(36, 1) = "�`��": Cnv2(36, 2) = "chu"
  Cnv2(37, 1) = "�a��": Cnv2(37, 2) = "dyu"
  Cnv2(38, 1) = "�`��": Cnv2(38, 2) = "cho"
  Cnv2(39, 1) = "�a��": Cnv2(39, 2) = "dyo"
  Cnv2(40, 1) = "�c�@": Cnv2(40, 2) = "tsa"
  Cnv2(41, 1) = "�c�B": Cnv2(41, 2) = "tsi"
  Cnv2(42, 1) = "�c�F": Cnv2(42, 2) = "tse"
  Cnv2(43, 1) = "�c�H": Cnv2(43, 2) = "tso"
  Cnv2(44, 1) = "�e�B": Cnv2(44, 2) = "thi"
  Cnv2(45, 1) = "�f�B": Cnv2(45, 2) = "dhi"
  Cnv2(46, 1) = "�e�F": Cnv2(46, 2) = "the"
  Cnv2(47, 1) = "�f�F": Cnv2(47, 2) = "dhe"
  Cnv2(48, 1) = "�e��": Cnv2(48, 2) = "tha"
  Cnv2(49, 1) = "�f��": Cnv2(49, 2) = "dha"
  Cnv2(50, 1) = "�e��": Cnv2(50, 2) = "thu"
  Cnv2(51, 1) = "�f��": Cnv2(51, 2) = "dhu"
  Cnv2(52, 1) = "�e��": Cnv2(52, 2) = "tho"
  Cnv2(53, 1) = "�f��": Cnv2(53, 2) = "dho"
  Cnv2(54, 1) = "�g�D": Cnv2(54, 2) = "twu"
  Cnv2(55, 1) = "�h�D": Cnv2(55, 2) = "dwu"
  Cnv2(56, 1) = "�j�B": Cnv2(56, 2) = "nyi"
  Cnv2(57, 1) = "�j�F": Cnv2(57, 2) = "nye"
  Cnv2(58, 1) = "�j��": Cnv2(58, 2) = "nya"
  Cnv2(59, 1) = "�j��": Cnv2(59, 2) = "nyu"
  Cnv2(60, 1) = "�j��": Cnv2(60, 2) = "nyo"
  Cnv2(61, 1) = "�q�B": Cnv2(61, 2) = "hyi"
  Cnv2(62, 1) = "�r�B": Cnv2(62, 2) = "byi"
  Cnv2(63, 1) = "�s�B": Cnv2(63, 2) = "pyi"
  Cnv2(64, 1) = "�q�F": Cnv2(64, 2) = "hye"
  Cnv2(65, 1) = "�r�F": Cnv2(65, 2) = "bye"
  Cnv2(66, 1) = "�s�F": Cnv2(66, 2) = "pye"
  Cnv2(67, 1) = "�q��": Cnv2(67, 2) = "hya"
  Cnv2(68, 1) = "�r��": Cnv2(68, 2) = "bya"
  Cnv2(69, 1) = "�s��": Cnv2(69, 2) = "pya"
  Cnv2(70, 1) = "�q��": Cnv2(70, 2) = "hyu"
  Cnv2(71, 1) = "�r��": Cnv2(71, 2) = "byu"
  Cnv2(72, 1) = "�s��": Cnv2(72, 2) = "pyu"
  Cnv2(73, 1) = "�q��": Cnv2(73, 2) = "hyo"
  Cnv2(74, 1) = "�r��": Cnv2(74, 2) = "byo"
  Cnv2(75, 1) = "�s��": Cnv2(75, 2) = "pyo"
  Cnv2(76, 1) = "�t�@": Cnv2(76, 2) = "fa"
  Cnv2(77, 1) = "�t�B": Cnv2(77, 2) = "fi"
  Cnv2(78, 1) = "�t�F": Cnv2(78, 2) = "fe"
  Cnv2(79, 1) = "�t�H": Cnv2(79, 2) = "fo"
  Cnv2(80, 1) = "�t��": Cnv2(80, 2) = "fya"
  Cnv2(81, 1) = "�t��": Cnv2(81, 2) = "fyu"
  Cnv2(82, 1) = "�t��": Cnv2(82, 2) = "fyo"
  Cnv2(83, 1) = "�~�B": Cnv2(83, 2) = "myi"
  Cnv2(84, 1) = "�~�F": Cnv2(84, 2) = "mye"
  Cnv2(85, 1) = "�~��": Cnv2(85, 2) = "mya"
  Cnv2(86, 1) = "�~��": Cnv2(86, 2) = "myu"
  Cnv2(87, 1) = "�~��": Cnv2(87, 2) = "myo"
  Cnv2(88, 1) = "���B": Cnv2(88, 2) = "ryi"
  Cnv2(89, 1) = "���F": Cnv2(89, 2) = "rye"
  Cnv2(90, 1) = "����": Cnv2(90, 2) = "rya"
  Cnv2(91, 1) = "����": Cnv2(91, 2) = "ryu"
  Cnv2(92, 1) = "����": Cnv2(92, 2) = "ryo"
  Cnv1(1, 1) = "�@": Cnv1(1, 2) = "la"
  Cnv1(2, 1) = "�A": Cnv1(2, 2) = "a"
  Cnv1(3, 1) = "�B": Cnv1(3, 2) = "li"
  Cnv1(4, 1) = "�C": Cnv1(4, 2) = "i"
  Cnv1(5, 1) = "�D": Cnv1(5, 2) = "lu"
  Cnv1(6, 1) = "�E": Cnv1(6, 2) = "u"
  Cnv1(7, 1) = "��": Cnv1(7, 2) = "vu"
  Cnv1(8, 1) = "�F": Cnv1(8, 2) = "le"
  Cnv1(9, 1) = "�G": Cnv1(9, 2) = "e"
  Cnv1(10, 1) = "�H": Cnv1(10, 2) = "lo"
  Cnv1(11, 1) = "�I": Cnv1(11, 2) = "o"
  Cnv1(12, 1) = "��": Cnv1(12, 2) = "lka"
  Cnv1(13, 1) = "�J": Cnv1(13, 2) = "ka"
  Cnv1(14, 1) = "�K": Cnv1(14, 2) = "ga"
  Cnv1(15, 1) = "�L": Cnv1(15, 2) = "ki"
  Cnv1(16, 1) = "�M": Cnv1(16, 2) = "gi"
  Cnv1(17, 1) = "�N": Cnv1(17, 2) = "ku"
  Cnv1(18, 1) = "�O": Cnv1(18, 2) = "gu"
  Cnv1(19, 1) = "��": Cnv1(19, 2) = "lke"
  Cnv1(20, 1) = "�P": Cnv1(20, 2) = "ke"
  Cnv1(21, 1) = "�Q": Cnv1(21, 2) = "ge"
  Cnv1(22, 1) = "�R": Cnv1(22, 2) = "ko"
  Cnv1(23, 1) = "�S": Cnv1(23, 2) = "go"
  Cnv1(24, 1) = "�T": Cnv1(24, 2) = "sa"
  Cnv1(25, 1) = "�U": Cnv1(25, 2) = "za"
  Cnv1(26, 1) = "�V": Cnv1(26, 2) = "shi"
  Cnv1(27, 1) = "�W": Cnv1(27, 2) = "ji"
  Cnv1(28, 1) = "�X": Cnv1(28, 2) = "su"
  Cnv1(29, 1) = "�Y": Cnv1(29, 2) = "zu"
  Cnv1(30, 1) = "�Z": Cnv1(30, 2) = "se"
  Cnv1(31, 1) = "�[": Cnv1(31, 2) = "ze"
  Cnv1(32, 1) = "�\": Cnv1(32, 2) = "so"
  Cnv1(33, 1) = "�]": Cnv1(33, 2) = "zo"
  Cnv1(34, 1) = "�^": Cnv1(34, 2) = "ta"
  Cnv1(35, 1) = "�_": Cnv1(35, 2) = "da"
  Cnv1(36, 1) = "�`": Cnv1(36, 2) = "chi"
  Cnv1(37, 1) = "�a": Cnv1(37, 2) = "di"
  Cnv1(38, 1) = "�c": Cnv1(38, 2) = "tsu"
  Cnv1(39, 1) = "�d": Cnv1(39, 2) = "du"
  Cnv1(40, 1) = "�e": Cnv1(40, 2) = "te"
  Cnv1(41, 1) = "�f": Cnv1(41, 2) = "de"
  Cnv1(42, 1) = "�g": Cnv1(42, 2) = "to"
  Cnv1(43, 1) = "�h": Cnv1(43, 2) = "do"
  Cnv1(44, 1) = "�i": Cnv1(44, 2) = "na"
  Cnv1(45, 1) = "�j": Cnv1(45, 2) = "ni"
  Cnv1(46, 1) = "�k": Cnv1(46, 2) = "nu"
  Cnv1(47, 1) = "�l": Cnv1(47, 2) = "ne"
  Cnv1(48, 1) = "�m": Cnv1(48, 2) = "no"
  Cnv1(49, 1) = "�n": Cnv1(49, 2) = "ha"
  Cnv1(50, 1) = "�o": Cnv1(50, 2) = "ba"
  Cnv1(51, 1) = "�p": Cnv1(51, 2) = "pa"
  Cnv1(52, 1) = "�q": Cnv1(52, 2) = "hi"
  Cnv1(53, 1) = "�r": Cnv1(53, 2) = "bi"
  Cnv1(54, 1) = "�s": Cnv1(54, 2) = "pi"
  Cnv1(55, 1) = "�t": Cnv1(55, 2) = "fu"
  Cnv1(56, 1) = "�u": Cnv1(56, 2) = "bu"
  Cnv1(57, 1) = "�v": Cnv1(57, 2) = "pu"
  Cnv1(58, 1) = "�w": Cnv1(58, 2) = "he"
  Cnv1(59, 1) = "�x": Cnv1(59, 2) = "be"
  Cnv1(60, 1) = "�y": Cnv1(60, 2) = "pe"
  Cnv1(61, 1) = "�z": Cnv1(61, 2) = "ho"
  Cnv1(62, 1) = "�{": Cnv1(62, 2) = "bo"
  Cnv1(63, 1) = "�|": Cnv1(63, 2) = "po"
  Cnv1(64, 1) = "�}": Cnv1(64, 2) = "ma"
  Cnv1(65, 1) = "�~": Cnv1(65, 2) = "mi"
  Cnv1(66, 1) = "��": Cnv1(66, 2) = "mu"
  Cnv1(67, 1) = "��": Cnv1(67, 2) = "me"
  Cnv1(68, 1) = "��": Cnv1(68, 2) = "mo"
  Cnv1(69, 1) = "��": Cnv1(69, 2) = "lya"
  Cnv1(70, 1) = "��": Cnv1(70, 2) = "ya"
  Cnv1(71, 1) = "��": Cnv1(71, 2) = "lyu"
  Cnv1(72, 1) = "��": Cnv1(72, 2) = "yu"
  Cnv1(73, 1) = "��": Cnv1(73, 2) = "lyo"
  Cnv1(74, 1) = "��": Cnv1(74, 2) = "yo"
  Cnv1(75, 1) = "��": Cnv1(75, 2) = "ra"
  Cnv1(76, 1) = "��": Cnv1(76, 2) = "ri"
  Cnv1(77, 1) = "��": Cnv1(77, 2) = "ru"
  Cnv1(78, 1) = "��": Cnv1(78, 2) = "re"
  Cnv1(79, 1) = "��": Cnv1(79, 2) = "ro"
  Cnv1(80, 1) = "��": Cnv1(80, 2) = "lwa"
  Cnv1(81, 1) = "��": Cnv1(81, 2) = "wa"
  Cnv1(82, 1) = "��": Cnv1(82, 2) = "wa"
  Cnv1(83, 1) = "��": Cnv1(83, 2) = "wyi"
  Cnv1(84, 1) = "��": Cnv1(84, 2) = "wye"
  Cnv1(85, 1) = "��": Cnv1(85, 2) = "wo"
  Cnv1(86, 1) = "��": Cnv1(86, 2) = "nn"
  Cnv1(87, 1) = "�[": Cnv1(87, 2) = "-"
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
    ElseIf Mid(kana, i, 1) = "�b" Then
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
'���[�}����S�p�J�i�ɕϊ�
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
  kanatbl(1) = Array("�A", "�C", "�E", "�G", "�I")
  kanatbl(2) = Array("�o", "�r", "�u", "�x", "�{")
  kanatbl(3) = Array("�_", "�a", "�d", "�f", "�h")
  kanatbl(4) = Array("�t�@", "�t�B", "�t", "�t�F", "�t�H")
  kanatbl(5) = Array("�K", "�M", "�O", "�Q", "�S")
  kanatbl(6) = Array("�n", "�q", "�t", "�w", "�z")
  kanatbl(7) = Array("�W��", "�W", "�W��", "�W�F", "�W��")
  kanatbl(8) = Array("�J", "�L", "�N", "�P", "�R")
  kanatbl(9) = Array("�@", "�B", "�D", "�F", "�H")
  kanatbl(10) = Array("�}", "�~", "��", "��", "��")
  kanatbl(11) = Array("�i", "�j", "�k", "�l", "�m")
  kanatbl(12) = Array("�p", "�s", "�v", "�y", "�|")
  kanatbl(13) = Array("��", "��", "��", "��", "��")
  kanatbl(14) = Array("�T", "�V", "�X", "�Z", "�\")
  kanatbl(15) = Array("�^", "�`", "�c", "�e", "�g")
  kanatbl(16) = Array("���@", "���B", "��", "���F", "���H")
  kanatbl(17) = Array("��", "�E�B", "�E", "�E�F", "��")
  kanatbl(18) = Array("�@", "�B", "�D", "�F", "�H")
  kanatbl(19) = Array("��", "�C", "��", "�C�F", "��")
  kanatbl(20) = Array("�U", "�W", "�Y", "�[", "�]")
  kanatbl(21) = Array("�r��", "�r�B", "�r��", "�r�F", "�r��")
  kanatbl(22) = Array("�`��", "�`", "�`��", "�`�F", "�`��")
  kanatbl(23) = Array("�`��", "�`�B", "�`��", "�`�F", "�`��")
  kanatbl(24) = Array("�f��", "�f�B", "�f��", "�f�F", "�f��")
  kanatbl(25) = Array("dwa", "dwi", "�h�D", "dwe", "dwo")
  kanatbl(26) = Array("�a��", "�a�B", "�a��", "�a�F", "�a��")
  kanatbl(27) = Array("�t��", "�t�B", "�t��", "�t�F", "�t��")
  kanatbl(28) = Array("�O�@", "gwi", "gwu", "gwe", "gwo")
  kanatbl(29) = Array("�M��", "�M�B", "�M��", "�M�F", "�M��")
  kanatbl(30) = Array("�q��", "�q�B", "�q��", "�q�F", "�q��")
  kanatbl(31) = Array("�W��", "�W�B", "�W��", "�W�F", "�W��")
  kanatbl(32) = Array("�N�@", "kwi", "kwu", "kwe", "kwo")
  kanatbl(33) = Array("�L��", "�L�B", "�L��", "�L�F", "�L��")
  kanatbl(34) = Array("��", "lki", "lku", "��", "lko")
  kanatbl(35) = Array("lta", "lti", "�b", "lte", "lto")
  kanatbl(36) = Array("��", "lwi", "lwu", "lwe", "lwo")
  kanatbl(37) = Array("��", "�B", "��", "�F", "��")
  kanatbl(38) = Array("�~��", "�~�B", "�~��", "�~�F", "�~��")
  kanatbl(39) = Array("���A", "���C", "���E", "���G", "���I")
  kanatbl(40) = Array("�j��", "�j�B", "�j��", "�j�F", "�j��")
  kanatbl(41) = Array("�s��", "�s�B", "�s��", "�s�F", "�s��")
  kanatbl(42) = Array("����", "���B", "����", "���F", "����")
  kanatbl(43) = Array("�V��", "�V", "�V��", "�V�F", "�V��")
  kanatbl(44) = Array("�V��", "�V�B", "�V��", "�V�F", "�V��")
  kanatbl(45) = Array("�e��", "�e�B", "�e��", "�e�F", "�e��")
  kanatbl(46) = Array("�c�@", "�c�B", "�c", "�c�F", "�c�H")
  kanatbl(47) = Array("twa", "twi", "�g�D", "twe", "two")
  kanatbl(48) = Array("�`��", "�`�B", "�`��", "�`�F", "�`��")
  kanatbl(49) = Array("wya", "��", "wyu", "��", "wyo")
  kanatbl(50) = Array("��", "xki", "xku", "��", "xko")
  kanatbl(51) = Array("xta", "xti", "�b", "xte", "xto")
  kanatbl(52) = Array("��", "xwi", "xwu", "xwe", "xwo")
  kanatbl(53) = Array("��", "�B", "��", "�F", "��")
  kanatbl(54) = Array("�W��", "�W�B", "�W��", "�W�F", "�W��")
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
           IIf(hiragana, Replace(StrConv(kanatbl(index)(k - 1), 36), "��", "���J"), _
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
            retStr = retStr & IIf(hiragana, "��", "�b")
          Case "nn"
            retStr = retStr & IIf(hiragana, "��", "��")
          Case "-"
            retStr = retStr & IIf(hiragana, "�[", "�[")
          Case "n"
            retStr = retStr & IIf(hiragana, "��", "��")
          Case Else
            retStr = retStr & Pre
        End Select
        Pre = Mid(roma, i, IIf(j2, 2, 1))
        i = i + 1 + IIf(j2, 1, 0)
        If Pre = "lt" Or Pre = "xt" Then
          If Mid(roma, i, 2) = "su" Then
            retStr = retStr + IIf(hiragana, "��", "�b")
            i = i + 2
            Pre = ""
          End If
        End If
      End If
    Else
      retStr = retStr + IIf(Pre = "nn" Or Pre = "n", IIf(hiragana, "��", "��"), Pre)
      If Mid(roma, i, 1) <> "'" Then
         retStr = retStr + Mid(roma, i, 1)
      End If
      Pre = ""
      index = 1
      i = i + 1
    End If
  Loop
  roma2kana = retStr & IIf(Pre = "nn" Or Pre = "n" Or Pre = "n'", IIf(hiragana, "��", "��"), Pre)
End Function



