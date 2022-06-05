Attribute VB_Name = "hebon"
Option Explicit
Function hebon(ByVal kana As String) As String
    Dim temp_data As String
    Dim temp_data2 As String
    Dim temp_data3 As String
    Dim temp_data4 As String
    Dim temp_data5 As String
    Dim temp_data6 As String
    Dim roma_data As String
    Dim result_data As String
    Dim sub_str2 As String
    Dim map As Dictionary
    Set map = initMap
    Dim i As Long
    For i = 1 To Len(kana)
        temp_data = Mid(kana, i, 1)
        temp_data2 = Mid(kana, i, 2)
        If Not map.Exists(temp_data2) Then
            If Not map.Exists(temp_data) Then
                roma_data = roma_data + temp_data
            Else
                roma_data = roma_data + map(temp_data)
            End If
        Else
            i = i + 1
            roma_data = roma_data + map(temp_data2)
        End If
    Next
    
    Dim k As Long
    For k = 1 To Len(roma_data)
        temp_data = Mid(roma_data, k, 1)
        temp_data2 = Mid(roma_data, k, 2)
        temp_data3 = Mid(roma_data, k, 3)
        temp_data4 = Mid(roma_data, k, 4)
        temp_data5 = Mid(roma_data, k, 5)
        temp_data6 = Mid(roma_data, k, 6)
        sub_str2 = Mid(temp_data2, 1, 2)
        If temp_data4 = "noue" Then
            k = k + 3
            temp_data = "noue"
        ElseIf temp_data6 = "touchi" Then
            k = k + 5
            temp_data = "touchi"
        Else
            
            Dim re As Object
            Set re = CreateObject("VBScript.RegExp")
            With re
                .Pattern = "/[a-z]/gi"       ''検索パターンを設定
                .IgnoreCase = True          ''大文字と小文字を区別しない
                .Global = True
            End With
            If temp_data2 = "uu" Or temp_data2 = "ee" Or temp_data2 = "ou" Or temp_data2 = "oo" Then
                k = k + 1
            ElseIf temp_data = "っ" Or temp_data = "ッ" Then
                If temp_data3 = "っch" Or temp_data3 = "ッch" Then
                    temp_data = "t"
                ElseIf re.test(sub_str2) Then
                    temp_data = sub_str2
                Else
                    temp_data = "tsu"
                End If
            End If
        End If
        result_data = result_data + temp_data
    Next
    
    hebon = result_data
End Function
Function initMap() As Dictionary
Static map As Dictionary
If Not map Is Nothing Then
    Set initMap = map
    Exit Function
End If
Set map = New Dictionary
With map
    .Add "あ", "a"
    .Add "い", "i"
    .Add "う", "u"
    .Add "え", "e"
    .Add "お", "o"
    .Add "か", "ka"
    .Add "き", "ki"
    .Add "く", "ku"
    .Add "け", "ke"
    .Add "こ", "ko"
    .Add "さ", "sa"
    .Add "し", "shi"
    .Add "す", "su"
    .Add "せ", "se"
    .Add "そ", "so"
    .Add "た", "ta"
    .Add "ち", "chi"
    .Add "つ", "tsu"
    .Add "て", "te"
    .Add "と", "to"
    .Add "な", "na"
    .Add "に", "ni"
    .Add "ぬ", "nu"
    .Add "ね", "ne"
    .Add "の", "no"
    .Add "は", "ha"
    .Add "ひ", "hi"
    .Add "ふ", "fu"
    .Add "へ", "he"
    .Add "ほ", "ho"
    .Add "ま", "ma"
    .Add "み", "mi"
    .Add "む", "mu"
    .Add "め", "me"
    .Add "も", "mo"
    .Add "や", "ya"
    .Add "ゆ", "yu"
    .Add "よ", "yo"
    .Add "ら", "ra"
    .Add "り", "ri"
    .Add "る", "ru"
    .Add "れ", "re"
    .Add "ろ", "ro"
    .Add "わ", "wa"
    .Add "ゐ", "i"
    .Add "ゑ", "e"
    .Add "を", "o"
    .Add "ん", "n"
    .Add "ぁ", "a"
    .Add "ぃ", "i"
    .Add "ぅ", "u"
    .Add "ぇ", "e"
    .Add "ぉ", "o"
    .Add "が", "ga"
    .Add "ぎ", "gi"
    .Add "ぐ", "gu"
    .Add "げ", "ge"
    .Add "ご", "go"
    .Add "ざ", "za"
    .Add "じ", "ji"
    .Add "ず", "zu"
    .Add "ぜ", "ze"
    .Add "ぞ", "zo"
    .Add "だ", "da"
    .Add "ぢ", "ji"
    .Add "づ", "zu"
    .Add "で", "de"
    .Add "ど", "do"
    .Add "ば", "ba"
    .Add "び", "bi"
    .Add "ぶ", "bu"
    .Add "べ", "be"
    .Add "ぼ", "bo"
    .Add "ぱ", "pa"
    .Add "ぴ", "pi"
    .Add "ぷ", "pu"
    .Add "ぺ", "pe"
    .Add "ぽ", "po"
    .Add "きゃ", "kya"
    .Add "きゅ", "kyu"
    .Add "きょ", "kyo"
    .Add "しゃ", "sha"
    .Add "しゅ", "shu"
    .Add "しょ", "sho"
    .Add "ちゃ", "cha"
    .Add "ちゅ", "chu"
    .Add "ちょ", "cho"
    .Add "ちぇ", "che"
    .Add "にゃ", "nya"
    .Add "にゅ", "nyu"
    .Add "にょ", "nyo"
    .Add "ひゃ", "hya"
    .Add "ひゅ", "hyu"
    .Add "ひょ", "hyo"
    .Add "みゃ", "mya"
    .Add "みゅ", "myu"
    .Add "みょ", "myo"
    .Add "りゃ", "rya"
    .Add "りゅ", "ryu"
    .Add "りょ", "ryo"
    .Add "ぎゃ", "gya"
    .Add "ぎゅ", "gyu"
    .Add "ぎょ", "gyo"
    .Add "じゃ", "ja"
    .Add "じゅ", "ju"
    .Add "じょ", "jo"
    .Add "びゃ", "bya"
    .Add "びゅ", "byu"
    .Add "びょ", "byo"
    .Add "ぴゃ", "pya"
    .Add "ぴゅ", "pyu"
    .Add "ぴょ", "pyo"
    .Add "ア", "a"
    .Add "イ", "i"
    .Add "ウ", "u"
    .Add "エ", "e"
    .Add "オ", "o"
    .Add "カ", "ka"
    .Add "キ", "ki"
    .Add "ク", "ku"
    .Add "ケ", "ke"
    .Add "コ", "ko"
    .Add "サ", "sa"
    .Add "シ", "shi"
    .Add "ス", "su"
    .Add "セ", "se"
    .Add "ソ", "so"
    .Add "タ", "ta"
    .Add "チ", "chi"
    .Add "ツ", "tsu"
    .Add "テ", "te"
    .Add "ト", "to"
    .Add "ナ", "na"
    .Add "ニ", "ni"
    .Add "ヌ", "nu"
    .Add "ネ", "ne"
    .Add "ノ", "no"
    .Add "ハ", "ha"
    .Add "ヒ", "hi"
    .Add "フ", "fu"
    .Add "ヘ", "he"
    .Add "ホ", "ho"
    .Add "マ", "ma"
    .Add "ミ", "mi"
    .Add "ム", "mu"
    .Add "メ", "me"
    .Add "モ", "mo"
    .Add "ヤ", "ya"
    .Add "ユ", "yu"
    .Add "ヨ", "yo"
    .Add "ラ", "ra"
    .Add "リ", "ri"
    .Add "ル", "ru"
    .Add "レ", "re"
    .Add "ロ", "ro"
    .Add "ワ", "wa"
    .Add "ヰ", "i"
    .Add "ヱ", "e"
    .Add "ヲ", "o"
    .Add "ン", "n"
    .Add "ァ", "a"
    .Add "ィ", "i"
    .Add "ゥ", "u"
    .Add "ェ", "e"
    .Add "ォ", "o"
    .Add "ガ", "ga"
    .Add "ギ", "gi"
    .Add "グ", "gu"
    .Add "ゲ", "ge"
    .Add "ゴ", "go"
    .Add "ザ", "za"
    .Add "ジ", "ji"
    .Add "ズ", "zu"
    .Add "ゼ", "ze"
    .Add "ゾ", "zo"
    .Add "ダ", "da"
    .Add "ヂ", "ji"
    .Add "ヅ", "zu"
    .Add "デ", "de"
    .Add "ド", "do"
    .Add "バ", "ba"
    .Add "ビ", "bi"
    .Add "ブ", "bu"
    .Add "ベ", "be"
    .Add "ボ", "bo"
    .Add "パ", "pa"
    .Add "ピ", "pi"
    .Add "プ", "pu"
    .Add "ペ", "pe"
    .Add "ポ", "po"
    .Add "キャ", "kya"
    .Add "キュ", "kyu"
    .Add "キョ", "kyo"
    .Add "シャ", "sha"
    .Add "シュ", "shu"
    .Add "ショ", "sho"
    .Add "チャ", "cha"
    .Add "チュ", "chu"
    .Add "チョ", "cho"
    .Add "ニャ", "nya"
    .Add "ニュ", "nyu"
    .Add "ニョ", "nyo"
    .Add "ヒャ", "hya"
    .Add "ヒュ", "hyu"
    .Add "ヒョ", "hyo"
    .Add "ミャ", "mya"
    .Add "ミュ", "myu"
    .Add "ミョ", "myo"
    .Add "リャ", "rya"
    .Add "リュ", "ryu"
    .Add "リョ", "ryo"
    .Add "ギャ", "gya"
    .Add "ギュ", "gyu"
    .Add "ギョ", "gyo"
    .Add "ジャ", "ja"
    .Add "ジュ", "ju"
    .Add "ジョ", "jo"
    .Add "ビャ", "bya"
    .Add "ビュ", "byu"
    .Add "ビョ", "byo"
    .Add "ピャ", "pya"
    .Add "ピュ", "pyu"
    .Add "ピョ", "pyo"
    .Add "ジェ", "jie"
    .Add "チェ", "chie"
    .Add "ティ", "tei"
    .Add "ディ", "dei"
    .Add "デュ", "deyu"
    .Add "ファ", "fua"
    .Add "フィ", "fui"
    .Add "フェ", "fue"
    .Add "フォ", "fuo"
    .Add "ヴァ", "bua"
    .Add "ヴィ", "bui"
    .Add "ヴ", "bu"
    .Add "ヴェ", "bue"
    .Add "ヴォ", "buo"
    .Add "ー", ""

End With
Set initMap = map
End Function
