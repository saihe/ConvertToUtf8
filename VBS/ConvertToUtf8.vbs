' 宣言
Dim fso
Dim istm
Dim ostm
Dim tmp
Dim bin
Dim dir
Dim idir
Dim odir
Dim ifile
Dim ofile
Dim record

' 初期化
Set fso = CreateObject("Scripting.FileSystemObject")
Set istm = CreateObject("ADODB.Stream")
Set ostm = CreateObject("ADODB.Stream")
Set tmp = CreateObject("ADODB.Stream")
dir = fso.getParentFolderName(WScript.ScriptFullName)
Set idir = fso.GetFolder(dir + "\input")
Set odir = fso.GetFolder(dir + "\output")

' パラメータ設定
istm.Type = 2	'1: バイナリ, 2: テキスト
istm.Charset = "UTF-8"
ostm.Type = 1
tmp.Type = 2
tmp.Charset = "UTF-8"

' 読み書き
For Each ifile in idir.Files
	istm.Open
	istm.LoadFromFile ifile
	tmp.Open
	
	' 読み込み
	Do Until istm.EOS
		Dim line
		line = istm.ReadText(adReadLine)	'-1:全行読み, -2:1行読み
		tmp.WriteText line, adWriteLine	'0:文字列の書き込み, 1:文字列＋改行を書き込み
		msgbox line
	Loop
	
	' バイナリモードにする
	tmp.Position = 0
	tmp.Type = 1
	tmp.Psition = 3	'先頭3倍と = BOMをスキップする
	
	' バイナリデータ読み込み
	bin = tmp.Read	' バイナリモードでないとReadできない
	
	' 書き出しファイルの指定 (読み込んだバイナリデータをバイナリデータとしてファイルに出力する)	ostm.Open
	ostm.Write(bin)
	
	' 書き出しファイルの保存
	ofile = odir + "\" + ifile.name
	ostm.SaveToFile ofile, 2	' 1:指定ファイルがなければ新規作成, 2:ファイルがある場合は上書き
	
	istm.Close
	tmp.Close
	ostm.Close
Next
